<#
.SYNOPSIS
    Acquire a token using MSAL.NET library.
.DESCRIPTION
    This command will acquire OAuth tokens for both public and confidential clients. Public clients authentication can be interactive, integrated Windows auth, or silent (aka refresh token authentication).
.EXAMPLE
    PS C:\>Get-MsalToken -ClientId '00000000-0000-0000-0000-000000000000' -Scope 'https://graph.microsoft.com/User.Read','https://graph.microsoft.com/Files.ReadWrite'
    Get AccessToken (with MS Graph permissions User.Read and Files.ReadWrite) and IdToken using client id from application registration (public client).
.EXAMPLE
    PS C:\>Get-MsalToken -ClientId '00000000-0000-0000-0000-000000000000' -TenantId '00000000-0000-0000-0000-000000000000' -Interactive -Scope 'https://graph.microsoft.com/User.Read' -LoginHint user@domain.com
    Force interactive authentication to get AccessToken (with MS Graph permissions User.Read) and IdToken for specific Azure AD tenant and UPN using client id from application registration (public client).
.EXAMPLE
    PS C:\>Get-MsalToken -ClientId '00000000-0000-0000-0000-000000000000' -ClientSecret (ConvertTo-SecureString 'SuperSecretString' -AsPlainText -Force) -TenantId '00000000-0000-0000-0000-000000000000' -Scope 'https://graph.microsoft.com/.default'
    Get AccessToken (with MS Graph permissions .Default) and IdToken for specific Azure AD tenant using client id and secret from application registration (confidential client).
.EXAMPLE
    PS C:\>$ClientCertificate = Get-Item Cert:\CurrentUser\My\0000000000000000000000000000000000000000
    PS C:\>$MsalClientApplication = Get-MsalClientApplication -ClientId '00000000-0000-0000-0000-000000000000' -ClientCertificate $ClientCertificate -TenantId '00000000-0000-0000-0000-000000000000'
    PS C:\>$MsalClientApplication | Get-MsalToken -Scope 'https://graph.microsoft.com/.default'
    Pipe in confidential client options object to get a confidential client application using a client certificate and target a specific tenant.
#>
function Get-MsalToken {
    [CmdletBinding(DefaultParameterSetName = 'PublicClient')]
    [OutputType([Microsoft.Identity.Client.AuthenticationResult])]
    param
    (
        # Identifier of the client requesting the token.
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-Interactive', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-IntegratedWindowsAuth', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-Silent', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-UsernamePassword', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-DeviceCode', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret-AuthorizationCode', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret-OnBehalfOf', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate-AuthorizationCode', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate-OnBehalfOf', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $ClientId,

        # Secure secret of the client requesting the token.
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret-AuthorizationCode', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret-OnBehalfOf', ValueFromPipelineByPropertyName = $true)]
        [securestring] $ClientSecret,

        # Client assertion certificate of the client requesting the token.
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate-AuthorizationCode', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate-OnBehalfOf', ValueFromPipelineByPropertyName = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $ClientCertificate,

        # Specifies if the x5c claim (public key of the certificate) should be sent to the STS.
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject')]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate')]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate-AuthorizationCode')]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate-OnBehalfOf')]
        [switch] $SendX5C,

        # The authorization code received from service authorization endpoint.
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject')]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret-AuthorizationCode')]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate-AuthorizationCode')]
        [string] $AuthorizationCode,

        # Assertion representing the user.
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret-OnBehalfOf', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate-OnBehalfOf', ValueFromPipelineByPropertyName = $true)]
        [string] $UserAssertion,

        # Type of the assertion representing the user.
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret-OnBehalfOf', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate-OnBehalfOf', ValueFromPipelineByPropertyName = $true)]
        [string] $UserAssertionType,

        # Address to return to upon receiving a response from the authority.
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-IntegratedWindowsAuth', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Silent', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-UsernamePassword', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-DeviceCode', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret-AuthorizationCode', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret-OnBehalfOf', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate-AuthorizationCode', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate-OnBehalfOf', ValueFromPipelineByPropertyName = $true)]
        [uri] $RedirectUri,

        # Instance of Azure Cloud
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.Identity.Client.AzureCloudInstance] $AzureCloudInstance,

        # Tenant identifier of the authority to issue token. It can also contain the value "consumers" or "organizations".
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string] $TenantId,

        # Address of the authority to issue token.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [uri] $Authority,

        # Use Platform Authentication Broker
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [switch] $AuthenticationBroker,

        # Public client application
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-InputObject', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.Identity.Client.IPublicClientApplication] $PublicClientApplication,

        # Confidential client application
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClient-InputObject', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.Identity.Client.IConfidentialClientApplication] $ConfidentialClientApplication,

        # Interactive request to acquire a token for the specified scopes.
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-Interactive')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject')]
        [switch] $Interactive,

        # Non-interactive request to acquire a security token for the signed-in user in Windows, via Integrated Windows Authentication.
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-IntegratedWindowsAuth')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject')]
        [switch] $IntegratedWindowsAuth,

        # Attempts to acquire an access token from the user token cache.
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-Silent')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject')]
        [switch] $Silent,

        # Acquires a security token on a device without a Web browser, by letting the user authenticate on another device.
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-DeviceCode')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject')]
        [switch] $DeviceCode,

        # Array of scopes requested for resource
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string[]] $Scopes = 'https://graph.microsoft.com/.default',

        # Array of scopes for which a developer can request consent upfront.
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [string[]] $ExtraScopesToConsent,

        # Identifier of the user. Generally a UPN.
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-IntegratedWindowsAuth', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Silent', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [string] $LoginHint,

        # Specifies the what the interactive experience is for the user. To force an interactive authentication, use the -Interactive switch.
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [ArgumentCompleter( {
                param ( $commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters )
                [Microsoft.Identity.Client.Prompt].DeclaredFields | Where-Object { $_.IsPublic -eq $true -and $_.IsStatic -eq $true -and $_.Name -like "$wordToComplete*" } | Select-Object -ExpandProperty Name
            })]
        [string] $Prompt,

        # Identifier of the user with associated password.
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-UsernamePassword', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [pscredential]
        [System.Management.Automation.Credential()]
        $UserCredential,

        # Correlation id to be used in the authentication request.
        [Parameter(Mandatory = $false)]
        [guid] $CorrelationId,

        # This parameter will be appended as is to the query string in the HTTP authentication request to the authority.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [hashtable] $ExtraQueryParameters,

        # Modifies the token acquisition request so that the acquired token is a Proof of Possession token (PoP), rather than a Bearer token.
        [Parameter(Mandatory = $false)]
        [System.Net.Http.HttpRequestMessage] $ProofOfPossession,

        # Ignore any access token in the user token cache and attempt to acquire new access token using the refresh token for the account if one is available.
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Silent')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject')]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret')]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate')]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject')]
        [switch] $ForceRefresh,

        # Specifies if the public client application should used an embedded web browser or the system default browser
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-Interactive', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [switch] $UseEmbeddedWebView,

        # Specifies the timeout threshold for MSAL.net operations.
        [Parameter(Mandatory = $false)]
        [timespan] $Timeout
    )

    begin {
        function CheckForMissingScopes([Microsoft.Identity.Client.AuthenticationResult]$AuthenticationResult, [string[]]$Scopes) {
            foreach ($Scope in $Scopes) {
                if ($AuthenticationResult.Scopes -notcontains $Scope) { return $true }
            }
            return $false
        }

        function Coalesce([psobject[]]$objects) { foreach ($object in $objects) { if ($object -notin $null, [string]::Empty) { return $object } } return $null }
    }

    process {
        switch -Wildcard ($PSCmdlet.ParameterSetName) {
            "PublicClient-InputObject" {
                [Microsoft.Identity.Client.IPublicClientApplication] $ClientApplication = $PublicClientApplication
                break
            }
            "ConfidentialClient-InputObject" {
                [Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplication
                break
            }
            "PublicClient*" {
                $paramSelectMsalClientApplication = Select-PsBoundParameters $PSBoundParameters -CommandName Select-MsalClientApplication -CommandParameterSets "PublicClient"
                [Microsoft.Identity.Client.IPublicClientApplication] $PublicClientApplication = Select-MsalClientApplication @paramSelectMsalClientApplication
                [Microsoft.Identity.Client.IPublicClientApplication] $ClientApplication = $PublicClientApplication
                break
            }
            "ConfidentialClientSecret*" {
                $paramSelectMsalClientApplication = Select-PsBoundParameters $PSBoundParameters -CommandName Select-MsalClientApplication -CommandParameterSets "ConfidentialClientSecret"
                [Microsoft.Identity.Client.IConfidentialClientApplication] $ConfidentialClientApplication = Select-MsalClientApplication @paramSelectMsalClientApplication
                [Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplication
                break
            }
            "ConfidentialClientCertificate*" {
                $paramSelectMsalClientApplication = Select-PsBoundParameters $PSBoundParameters -CommandName Select-MsalClientApplication -CommandParameterSets "ConfidentialClientCertificate"
                [Microsoft.Identity.Client.IConfidentialClientApplication] $ConfidentialClientApplication = Select-MsalClientApplication @paramSelectMsalClientApplication
                [Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplication
                break
            }
        }

        [Microsoft.Identity.Client.AuthenticationResult] $AuthenticationResult = $null
        switch -Wildcard ($PSCmdlet.ParameterSetName) {
            "PublicClient*" {
                if ($PSBoundParameters.ContainsKey("UserCredential") -and $UserCredential) {
                    $AquireTokenParameters = $PublicClientApplication.AcquireTokenByUsernamePassword($Scopes, $UserCredential.UserName, $UserCredential.Password)
                }
                elseif ($PSBoundParameters.ContainsKey("DeviceCode") -and $DeviceCode -or ($Interactive -and !$script:ModuleFeatureSupport.WebView1Support -and !$script:ModuleFeatureSupport.WebView2Support -and $PublicClientApplication.AppConfig.RedirectUri -ne 'http://localhost') -or ($Interactive -and !$script:ModuleFeatureSupport.WebView1Support -and $PublicClientApplication.AppConfig.RedirectUri -eq 'urn:ietf:wg:oauth:2.0:oob')) {
                    $AquireTokenParameters = $PublicClientApplication.AcquireTokenWithDeviceCode($Scopes, [DeviceCodeHelper]::GetDeviceCodeResultCallback())
                }
                elseif ($PSBoundParameters.ContainsKey("Interactive") -and $Interactive) {
                    $AquireTokenParameters = $PublicClientApplication.AcquireTokenInteractive($Scopes)
                    [IntPtr] $ParentWindow = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle
                    if ($ParentWindow -eq [System.IntPtr]::Zero -and [System.Environment]::OSVersion.Platform -eq 'Win32NT') {
                        $Win32Process = Get-CimInstance Win32_Process -Filter ("ProcessId = '{0}'" -f [System.Diagnostics.Process]::GetCurrentProcess().Id) -Verbose:$false
                        $ParentWindow = (Get-Process -Id $Win32Process.ParentProcessId).MainWindowHandle
                    }
                    if ($ParentWindow -ne [System.IntPtr]::Zero) { [void] $AquireTokenParameters.WithParentActivityOrWindow($ParentWindow) }
                    #if ($Account) { [void] $AquireTokenParameters.WithAccount($Account) }
                    if ($extraScopesToConsent) { [void] $AquireTokenParameters.WithExtraScopesToConsent($extraScopesToConsent) }
                    if ($LoginHint) { [void] $AquireTokenParameters.WithLoginHint($LoginHint) }
                    if ($Prompt) { [void] $AquireTokenParameters.WithPrompt([Microsoft.Identity.Client.Prompt]::$Prompt) }
                    if ($PSBoundParameters.ContainsKey('UseEmbeddedWebView')) { [void] $AquireTokenParameters.WithUseEmbeddedWebView($UseEmbeddedWebView) }
                    if (!$Timeout -and (($PSBoundParameters.ContainsKey('UseEmbeddedWebView') -and !$UseEmbeddedWebView) -or $PSVersionTable.PSEdition -eq 'Core')) {
                        $Timeout = New-TimeSpan -Minutes 2
                    }
                }
                elseif ($PSBoundParameters.ContainsKey("IntegratedWindowsAuth") -and $IntegratedWindowsAuth) {
                    $AquireTokenParameters = $PublicClientApplication.AcquireTokenByIntegratedWindowsAuth($Scopes)
                    if ($LoginHint) { [void] $AquireTokenParameters.WithUsername($LoginHint) }
                }
                elseif ($PSBoundParameters.ContainsKey("Silent") -and $Silent) {
                    if ($LoginHint) {
                        $AquireTokenParameters = $PublicClientApplication.AcquireTokenSilent($Scopes, $LoginHint)
                    }
                    else {
                        [Microsoft.Identity.Client.IAccount] $Account = $PublicClientApplication.GetAccountsAsync().GetAwaiter().GetResult() | Select-Object -First 1
                        $AquireTokenParameters = $PublicClientApplication.AcquireTokenSilent($Scopes, $Account)
                    }
                    if ($PSBoundParameters.ContainsKey('ForceRefresh')) { [void] $AquireTokenParameters.WithForceRefresh($ForceRefresh) }
                }
                else {
                    $paramGetMsalToken = Select-PsBoundParameters -NamedParameter $PSBoundParameters -CommandName 'Get-MsalToken' -CommandParameterSet 'PublicClient-InputObject' -ExcludeParameters 'PublicClientApplication'
                    ## Try Silent Authentication
                    Write-Verbose ('Attempting Silent Authentication to Application with ClientId [{0}]' -f $ClientApplication.ClientId)
                    try {
                        $AuthenticationResult = Get-MsalToken -Silent -PublicClientApplication $PublicClientApplication @paramGetMsalToken
                        ## Check for requested scopes
                        if (CheckForMissingScopes $AuthenticationResult $Scopes) {
                            $AuthenticationResult = Get-MsalToken -Interactive -PublicClientApplication $PublicClientApplication @paramGetMsalToken
                        }
                    }
                    catch [Microsoft.Identity.Client.MsalUiRequiredException] {
                        Write-Debug ('{0}: {1}' -f $_.Exception.GetType().Name, $_.Exception.Message)
                        ## Try Integrated Windows Authentication
                        Write-Verbose ('Attempting Integrated Windows Authentication to Application with ClientId [{0}]' -f $ClientApplication.ClientId)
                        try {
                            $AuthenticationResult = Get-MsalToken -IntegratedWindowsAuth -PublicClientApplication $PublicClientApplication @paramGetMsalToken
                            ## Check for requested scopes
                            if (CheckForMissingScopes $AuthenticationResult $Scopes) {
                                $AuthenticationResult = Get-MsalToken -Interactive -PublicClientApplication $PublicClientApplication @paramGetMsalToken
                            }
                        }
                        catch {
                            Write-Debug ('{0}: {1}' -f $_.Exception.GetType().Name, $_.Exception.Message)
                            ## Revert to Interactive Authentication
                            Write-Verbose ('Attempting Interactive Authentication to Application with ClientId [{0}]' -f $ClientApplication.ClientId)
                            $AuthenticationResult = Get-MsalToken -Interactive -PublicClientApplication $PublicClientApplication @paramGetMsalToken
                        }
                    }
                    break
                }
            }
            "ConfidentialClient*" {
                if ($PSBoundParameters.ContainsKey("AuthorizationCode")) {
                    $AquireTokenParameters = $ConfidentialClientApplication.AcquireTokenByAuthorizationCode($Scopes, $AuthorizationCode)
                }
                elseif ($PSBoundParameters.ContainsKey("UserAssertion")) {
                    if ($UserAssertionType) { [Microsoft.Identity.Client.UserAssertion] $UserAssertionObj = New-Object Microsoft.Identity.Client.UserAssertion -ArgumentList $UserAssertion, $UserAssertionType }
                    else { [Microsoft.Identity.Client.UserAssertion] $UserAssertionObj = New-Object Microsoft.Identity.Client.UserAssertion -ArgumentList $UserAssertion }
                    $AquireTokenParameters = $ConfidentialClientApplication.AcquireTokenOnBehalfOf($Scopes, $UserAssertionObj)
                }
                else {
                    $AquireTokenParameters = $ConfidentialClientApplication.AcquireTokenForClient($Scopes)
                    if ($PSBoundParameters.ContainsKey('ForceRefresh')) { [void] $AquireTokenParameters.WithForceRefresh($ForceRefresh) }
                }
                if ($SendX5C) { [void] $AquireTokenParameters.WithSendX5C($SendX5C) }
            }
            "*" {
                if ($AzureCloudInstance -and $TenantId) { [void] $AquireTokenParameters.WithAuthority($AzureCloudInstance, $TenantId) }
                elseif ($AzureCloudInstance) { [void] $AquireTokenParameters.WithAuthority($AzureCloudInstance, 'common') }
                elseif ($TenantId) { [void] $AquireTokenParameters.WithAuthority(('https://{0}' -f $ClientApplication.AppConfig.Authority.AuthorityInfo.Host), $TenantId) }
                if ($Authority) { [void] $AquireTokenParameters.WithAuthority($Authority.AbsoluteUri) }
                if ($CorrelationId) { [void] $AquireTokenParameters.WithCorrelationId($CorrelationId) }
                if ($ExtraQueryParameters) { [void] $AquireTokenParameters.WithExtraQueryParameters((ConvertTo-Dictionary $ExtraQueryParameters -KeyType ([string]) -ValueType ([string]))) }
                if ($ProofOfPossession) { [void] $AquireTokenParameters.WithProofOfPosession($ProofOfPossession) }
                Write-Debug ('Aquiring Token for Application with ClientId [{0}]' -f $ClientApplication.ClientId)
                if (!$Timeout) { $Timeout = [timespan]::Zero }

                ## Wait for async task to complete
                $tokenSource = New-Object System.Threading.CancellationTokenSource
                try {
                    #$AuthenticationResult = $AquireTokenParameters.ExecuteAsync().GetAwaiter().GetResult()
                    $taskAuthenticationResult = $AquireTokenParameters.ExecuteAsync($tokenSource.Token)
                    try {
                        $endTime = [datetime]::Now.Add($Timeout)
                        while (!$taskAuthenticationResult.IsCompleted) {
                            if ($Timeout -eq [timespan]::Zero -or [datetime]::Now -lt $endTime) {
                                Start-Sleep -Seconds 1
                            }
                            else {
                                $tokenSource.Cancel()
                                try { $taskAuthenticationResult.Wait() }
                                catch { }
                                Write-Error -Exception (New-Object System.TimeoutException) -Category ([System.Management.Automation.ErrorCategory]::OperationTimeout) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'GetMsalTokenFailureOperationTimeout' -TargetObject $AquireTokenParameters -ErrorAction Stop
                            }
                        }
                    }
                    finally {
                        if (!$taskAuthenticationResult.IsCompleted) {
                            Write-Debug ('Canceling Token Acquisition for Application with ClientId [{0}]' -f $ClientApplication.ClientId)
                            $tokenSource.Cancel()
                        }
                        $tokenSource.Dispose()
                    }

                    ## Parse task results
                    if ($taskAuthenticationResult.IsFaulted) {
                        Write-Error -Exception $taskAuthenticationResult.Exception -Category ([System.Management.Automation.ErrorCategory]::AuthenticationError) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'GetMsalTokenFailureAuthenticationError' -TargetObject $AquireTokenParameters -ErrorAction Stop
                    }
                    if ($taskAuthenticationResult.IsCanceled) {
                        Write-Error -Exception (New-Object System.Threading.Tasks.TaskCanceledException $taskAuthenticationResult) -Category ([System.Management.Automation.ErrorCategory]::OperationStopped) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'GetMsalTokenFailureOperationStopped' -TargetObject $AquireTokenParameters -ErrorAction Stop
                    }
                    else {
                        $AuthenticationResult = $taskAuthenticationResult.Result
                    }
                }
                catch {
                    Write-Error -Exception (Coalesce $_.Exception.InnerException,$_.Exception) -Category ([System.Management.Automation.ErrorCategory]::AuthenticationError) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'GetMsalTokenFailureAuthenticationError' -TargetObject $AquireTokenParameters -ErrorAction Stop
                }
                break
            }
        }

        return $AuthenticationResult
    }
}

# SIG # Begin signature block
# MIIjhgYJKoZIhvcNAQcCoIIjdzCCI3MCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBn3hFm+jQLrbIY
# yrTfo3gnwpxOgzZiKNnCcEa1g5V0UqCCDYEwggX/MIID56ADAgECAhMzAAACUosz
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
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVWzCCFVcCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAlKLM6r4lfM52wAAAAACUjAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgGTEkrp6r
# ApZDMjRgsKbkCjnAPr+m/0iRWXEBonFYbGIwQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQCOp9qqLqZ3KK/CJMl4IrhDdOXOvgtzFSXwtwp/hYZa
# PrRVDMuTwqqGEU9U/Junxl6hgJ/TNsPIlVGFB4sXDrNkuPj2cfXQP2FzfVmzSQsL
# tdCOJ6qWx/oU9kRlpopELvNLQRs4HeejV85uR6ROLh+4oD7TuYqq25E168P/RjPW
# wyScgFoGKp79Du3DDEBcbq5eegcvli06kEyabZCrPM4bwPz+KB6Go6y9CcxjSUno
# 8R5TsAytpczDCx3yuHbkt0zE01anvcjCZbcyJHk9qwtlBOPtCN9m5QDlsEQJkFWg
# dQ4XfcvFSyLKXpM6weTLjmc89UnLDkEU4LsEkiBa0UQfoYIS5TCCEuEGCisGAQQB
# gjcDAwExghLRMIISzQYJKoZIhvcNAQcCoIISvjCCEroCAQMxDzANBglghkgBZQME
# AgEFADCCAVEGCyqGSIb3DQEJEAEEoIIBQASCATwwggE4AgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEINqom+LDxnJIDVsGgn9q2pUxOQKxg4gNhPSxq2uA
# UuOrAgZhktYPxMIYEzIwMjExMTE5MDIzNzQ2Ljk1N1owBIACAfSggdCkgc0wgcox
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1p
# Y3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlvbnMxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
# RVNOOkU1QTYtRTI3Qy01OTJFMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloIIOPDCCBPEwggPZoAMCAQICEzMAAAFHnY/x5t4xg1kAAAAAAUcw
# DQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcN
# MjAxMTEyMTgyNTU1WhcNMjIwMjExMTgyNTU1WjCByjELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2Eg
# T3BlcmF0aW9uczEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046RTVBNi1FMjdDLTU5
# MkUxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCtBQNM6X32KFk/BJ8YaprfzEt6Lj34
# G+VLjzgfEgOGSVd1Mu7nCphK0K4oyPrzItgNRjB4gUiKq6GzgxdDHgZPgTEvm57z
# sascyGrybWkf3VVr8bqf2PIgGvwKDNEgVcygsEbuWwXz9Li6M7AOoD4TB8fl4ATm
# +L7b4+lYDUMJYMLzpiJzM745a0XHiriUaOpYWfkwO9Hz6uf+k2Hq7yGyguH8naPL
# MnYfmYIt2PXAwWVvG4MD4YbjXBVZ14ueh7YlqZTMua3n9kT1CZDsHvz+o58nsoam
# XRwRFOb7LDjVV++cZIZLO29usiI0H79tb3fSvh9tU7QC7CirNCBYagNJAgMBAAGj
# ggEbMIIBFzAdBgNVHQ4EFgQUtPjcb95koYZXGy9DPxN49dSCsLowHwYDVR0jBBgw
# FoAU1WM6XIoxkPNDe3xGG8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
# L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljVGltU3RhUENB
# XzIwMTAtMDctMDEuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0
# cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNUaW1TdGFQQ0FfMjAx
# MC0wNy0wMS5jcnQwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDCDAN
# BgkqhkiG9w0BAQsFAAOCAQEAUMQOyjV+ea2kEtXqD0cOfD2Z2PFUIy5kLkGU53RD
# GcfhlzIR9QlTgZLqTEhgLLuCSy6jcma+nPg7e5Xg1oqCZcZJRwtRPzS1F6/M6YR3
# 5H3brN0maVnPrmrQ91kkfsNqDTtuWDiAIBfkNEgCpQZCb4OV3HMu5L8eZzg5dUaJ
# 7XE+LBuphJSLFJtabxYt4fkCQxnTD2z50Y32ZuXiNmFFia7qVq+3Yc3mmW02+/KW
# H8P1HPiobJG8crGYgSEkxtkUXGdoutwGWW88KR9RRcM/4GKLqt2OQ8AWEQb7shgM
# 8pxNvu30TxejRApa4WAfOAejTG4+KzBm67XjVZ2IlXAPkjCCBnEwggRZoAMCAQIC
# CmEJgSoAAAAAAAIwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRp
# ZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTEwMDcwMTIxMzY1NVoXDTI1MDcwMTIx
# NDY1NVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCpHQ28dxGKOiDs/BOX9fp/aZRrdFQQ1aUKAIKF
# ++18aEssX8XD5WHCdrc+Zitb8BVTJwQxH0EbGpUdzgkTjnxhMFmxMEQP8WCIhFRD
# DNdNuDgIs0Ldk6zWczBXJoKjRQ3Q6vVHgc2/JGAyWGBG8lhHhjKEHnRhZ5FfgVSx
# z5NMksHEpl3RYRNuKMYa+YaAu99h/EbBJx0kZxJyGiGKr0tkiVBisV39dx898Fd1
# rL2KQk1AUdEPnAY+Z3/1ZsADlkR+79BL/W7lmsqxqPJ6Kgox8NpOBpG2iAg16Hgc
# sOmZzTznL0S6p/TcZL2kAcEgCZN4zfy8wMlEXV4WnAEFTyJNAgMBAAGjggHmMIIB
# 4jAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQU1WM6XIoxkPNDe3xGG8UzaFqF
# bVUwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1Ud
# EwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYD
# VR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwv
# cHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEB
# BE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwgaAGA1UdIAEB/wSBlTCB
# kjCBjwYJKwYBBAGCNy4DMIGBMD0GCCsGAQUFBwIBFjFodHRwOi8vd3d3Lm1pY3Jv
# c29mdC5jb20vUEtJL2RvY3MvQ1BTL2RlZmF1bHQuaHRtMEAGCCsGAQUFBwICMDQe
# MiAdAEwAZQBnAGEAbABfAFAAbwBsAGkAYwB5AF8AUwB0AGEAdABlAG0AZQBuAHQA
# LiAdMA0GCSqGSIb3DQEBCwUAA4ICAQAH5ohRDeLG4Jg/gXEDPZ2joSFvs+umzPUx
# vs8F4qn++ldtGTCzwsVmyWrf9efweL3HqJ4l4/m87WtUVwgrUYJEEvu5U4zM9GAS
# inbMQEBBm9xcF/9c+V4XNZgkVkt070IQyK+/f8Z/8jd9Wj8c8pl5SpFSAK84Dxf1
# L3mBZdmptWvkx872ynoAb0swRCQiPM/tA6WWj1kpvLb9BOFwnzJKJ/1Vry/+tuWO
# M7tiX5rbV0Dp8c6ZZpCM/2pif93FSguRJuI57BlKcWOdeyFtw5yjojz6f32WapB4
# pm3S4Zz5Hfw42JT0xqUKloakvZ4argRCg7i1gJsiOCC1JeVk7Pf0v35jWSUPei45
# V3aicaoGig+JFrphpxHLmtgOR5qAxdDNp9DvfYPw4TtxCd9ddJgiCGHasFAeb73x
# 4QDf5zEHpJM692VHeOj4qEir995yfmFrb3epgcunCaw5u+zGy9iCtHLNHfS4hQEe
# gPsbiSpUObJb2sgNVZl6h3M7COaYLeqN4DMuEin1wC9UJyH3yKxO2ii4sanblrKn
# QqLJzxlBTeCG+SqaoxFmMNO7dDJL32N79ZmKLxvHIa9Zta7cRDyXUHHXodLFVeNp
# 3lfB0d4wwP3M5k37Db9dT+mdHhk4L7zPWAUu7w2gUDXa7wknHNWzfjUeCLraNtvT
# X4/edIhJEqGCAs4wggI3AgEBMIH4oYHQpIHNMIHKMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBP
# cGVyYXRpb25zMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpFNUE2LUUyN0MtNTky
# RTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcG
# BSsOAwIaAxUAq6fBtEENocNASMqL03zGJS0wZd2ggYMwgYCkfjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIFAOVBSOswIhgPMjAy
# MTExMTkwNTQ5MzFaGA8yMDIxMTEyMDA1NDkzMVowdzA9BgorBgEEAYRZCgQBMS8w
# LTAKAgUA5UFI6wIBADAKAgEAAgIT2AIB/zAHAgEAAgIRXDAKAgUA5UKaawIBADA2
# BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIB
# AAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBAHgwX1oDX44P6r66HxWE0+lUZ+Z6N1lN
# iy5MCRU7aQS+ezpF+WB6gljGxIUr1yWis8eJgVeenQ/cIZDlZxGG+or4VCrXb/6Z
# LVkYmnAzyDs04NWSWHSYTlpdajyJ/Xox1Oh/BtNhr9QzuJ4uqNKcRLZxd2BdYCWf
# uET2BpGprVtsMYIDDTCCAwkCAQEwgZMwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
# IDIwMTACEzMAAAFHnY/x5t4xg1kAAAAAAUcwDQYJYIZIAWUDBAIBBQCgggFKMBoG
# CSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQxIgQgjt6Vv/w5
# JStgSUIbkrJTcmYwUe51N+AR2IWM14SFdz4wgfoGCyqGSIb3DQEJEAIvMYHqMIHn
# MIHkMIG9BCB72zwSA5TPugbIiZO/2H1hrisAVItwzDscb0WqihjphTCBmDCBgKR+
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAABR52P8ebeMYNZAAAA
# AAFHMCIEIO2OUZ+KmEoWcQ+E0bi5z6Uiaeo5Uqu3bPqKMEzGyKlDMA0GCSqGSIb3
# DQEBCwUABIIBAD4RcBMHbQ9rqCqytzAfmuCVvTqny3aXSICdqLiAGRNYaKP1Ra5J
# Hyh/e+mOltkLKiWGcPQ1+b4hh+9e5YyTqCgWXDtp0ay3sDT+qCFe+SBAfP9Kp8p1
# BcssBTBjqAkd0+U6h72+SLl4yZ4Ug9Hr2v6qu6lGrh5GTEUUBUSwi7qb0BXUUlLE
# i0+nDEU7QRNwkxetg98/qESPn0cZu0xk9xdpCsjmVfJCvRzDo2wqQnvlaXBLPOv3
# cl0jRX/Bg06hebyW/kBIwTUIhoI4azqsf4l412bxkM5Wh5Usl5iuR+LxWPNtszT2
# 6iSpXL+s2r+ZVt/fXO1E1cQhCy86pOu210g=
# SIG # End signature block
