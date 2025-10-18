<#
.SYNOPSIS
    Acquire a token using MSAL.NET library.
.DESCRIPTION
    This command will acquire OAuth tokens for both public and confidential clients. Public clients authentication can be interactive, Integrated Windows Authentication, or silent (aka refresh token authentication).
.EXAMPLE
    PS C:\>Get-MsalToken -ClientId '00000000-0000-0000-0000-000000000000' -Scope 'https://graph.microsoft.com/User.Read','https://graph.microsoft.com/Files.ReadWrite'
    Get AccessToken (with MS Graph permissions User.Read and Files.ReadWrite) and IdToken using client id from application registration (public client).
.EXAMPLE
    PS C:\>Get-MsalToken -ClientId '00000000-0000-0000-0000-000000000000' -TenantId '00000000-0000-0000-0000-000000000000' -Interactive -Scope 'https://graph.microsoft.com/User.Read' -LoginHint user@domain.com
    Force interactive authentication to get AccessToken (with MS Graph permissions User.Read) and IdToken for specific Entra ID tenant and UPN using client id from application registration (public client).
.EXAMPLE
    PS C:\>Get-MsalToken -ClientId '00000000-0000-0000-0000-000000000000' -ClientSecret (ConvertTo-SecureString 'SuperSecretString' -AsPlainText -Force) -TenantId '00000000-0000-0000-0000-000000000000' -Scope 'https://graph.microsoft.com/.default'
    Get AccessToken (with MS Graph permissions .Default) and IdToken for specific Entra ID tenant using client id and secret from application registration (confidential client).
.EXAMPLE
    PS C:\>$ClientCertificate = Get-Item -LiteralPath Cert:\CurrentUser\My\0000000000000000000000000000000000000000
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

    # BrowserRedirectError
    [uri] $BrowserRedirectError,

    # BrowserRedirectSuccess
    [uri] $BrowserRedirectSuccess,

    # HtmlMessageSuccess
    [string] $HtmlMessageSuccess,

    # HtmlMessageError
    [string] $HtmlMessageError,

    # Silent request to acquire a security token for the signed-in user in Windows, via Integrated Windows Authentication.
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

    $InteractiveAuthTopLevelParentWindow = $null
  }

  process {
    switch -Wildcard ($PSCmdlet.ParameterSetName) {
      'PublicClient-InputObject' {
        [Microsoft.Identity.Client.IPublicClientApplication] $ClientApplication = $PublicClientApplication
        break
      }
      'ConfidentialClient-InputObject' {
        [Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplication
        break
      }
      'PublicClient*' {
        $paramSelectMsalClientApplication = Select-PsBoundParameters $PSBoundParameters -CommandName Select-MsalClientApplication -CommandParameterSets 'PublicClient'
        [Microsoft.Identity.Client.IPublicClientApplication] $PublicClientApplication = Select-MsalClientApplication @paramSelectMsalClientApplication
        [Microsoft.Identity.Client.IPublicClientApplication] $ClientApplication = $PublicClientApplication
        break
      }
      'ConfidentialClientSecret*' {
        $paramSelectMsalClientApplication = Select-PsBoundParameters $PSBoundParameters -CommandName Select-MsalClientApplication -CommandParameterSets 'ConfidentialClientSecret'
        [Microsoft.Identity.Client.IConfidentialClientApplication] $ConfidentialClientApplication = Select-MsalClientApplication @paramSelectMsalClientApplication
        [Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplication
        break
      }
      'ConfidentialClientCertificate*' {
        $paramSelectMsalClientApplication = Select-PsBoundParameters $PSBoundParameters -CommandName Select-MsalClientApplication -CommandParameterSets 'ConfidentialClientCertificate'
        [Microsoft.Identity.Client.IConfidentialClientApplication] $ConfidentialClientApplication = Select-MsalClientApplication @paramSelectMsalClientApplication
        [Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplication
        break
      }
    }

    [Microsoft.Identity.Client.AuthenticationResult] $AuthenticationResult = $null
    switch -Wildcard ($PSCmdlet.ParameterSetName) {
      'PublicClient*' {
        if ($PSBoundParameters.ContainsKey('UserCredential') -and $UserCredential) {
          $AquireTokenParameters = $PublicClientApplication.AcquireTokenByUsernamePassword($Scopes, $UserCredential.UserName, $UserCredential.Password)
        } elseif (($PSBoundParameters.ContainsKey('DeviceCode') -and $DeviceCode) -or ($PSBoundParameters.ContainsKey('Interactive') -and $Interactive -and !$script:ModuleFeatureSupport.WebView1Support -and !$script:ModuleFeatureSupport.WebView2Support -and ([uri]$PublicClientApplication.AppConfig.RedirectUri).AbsoluteUri -ine 'http://localhost/') -or ($PSBoundParameters.ContainsKey('Interactive') -and $Interactive -and !$script:ModuleFeatureSupport.WebView1Support -and $PublicClientApplication.AppConfig.RedirectUri -ieq 'urn:ietf:wg:oauth:2.0:oob')) {
          $AquireTokenParameters = $PublicClientApplication.AcquireTokenWithDeviceCode($Scopes, [DeviceCodeHelper]::GetDeviceCodeResultCallback())
        } elseif ($PSBoundParameters.ContainsKey('Interactive') -and $Interactive) {
          $AquireTokenParameters = $PublicClientApplication.AcquireTokenInteractive($Scopes)

          if ((-not (Test-Path -LiteralPath 'variable:IsWindows')) -or $IsWindows) {
            [IntPtr] $ParentWindow = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle

            if ($ParentWindow -eq [System.IntPtr]::Zero) {
              Add-Type -AssemblyName PresentationCore, PresentationFramework, System.Windows.Forms

              $InteractiveAuthTopLevelParentWindow = New-Object System.Windows.Window -Property @{
                Width                 = 1
                Height                = 1
                WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
                ShowActivated         = $false
                Topmost               = $true
              }

              $InteractiveAuthTopLevelParentWindow.Show()
              $InteractiveAuthTopLevelParentWindow.Hide()

              [IntPtr] $ParentWindow = [System.Windows.Interop.WindowInteropHelper]::new($InteractiveAuthTopLevelParentWindow).Handle
            }

            if ($ParentWindow -ne [System.IntPtr]::Zero) { [void] $AquireTokenParameters.WithParentActivityOrWindow($ParentWindow) }
          } elseif ((Test-Path -LiteralPath 'variable:IsMacOS') -and $IsMacOS) {
            $objcCode = @'
using System;
using System.Runtime.InteropServices;

public class MacInterop {
    [DllImport("/usr/lib/libobjc.A.dylib", EntryPoint = "objc_getClass")]
    public static extern IntPtr objc_getClass(string className);

    [DllImport("/usr/lib/libobjc.A.dylib", EntryPoint = "sel_registerName")]
    public static extern IntPtr sel_registerName(string selectorName);

    [DllImport("/usr/lib/libobjc.A.dylib", EntryPoint = "objc_msgSend")]
    public static extern IntPtr objc_msgSend(IntPtr receiver, IntPtr selector);

    public static IntPtr GetMainWindowHandle() {
        IntPtr nsAppClass = objc_getClass("NSApplication");
        IntPtr sharedAppSel = sel_registerName("sharedApplication");
        IntPtr mainWindowSel = sel_registerName("mainWindow");

        IntPtr nsApp = objc_msgSend(nsAppClass, sharedAppSel);
        IntPtr mainWindow = objc_msgSend(nsApp, mainWindowSel);

        return mainWindow;
    }
}
'@

            Add-Type -TypeDefinition $objcCode -Language CSharp
            [IntPtr] $ParentWindow = [MacInterop]::GetMainWindowHandle()

            if ($ParentWindow -ne [System.IntPtr]::Zero) {
              [void] $AquireTokenParameters.WithParentActivityOrWindow($ParentWindow)
            }
          } elseif ((Test-Path -LiteralPath 'variable:IsLinux') -and $IsLinux) {
            try {
              $x11Code = @'
using System;
using System.Runtime.InteropServices;

public class X11Interop {
    [DllImport("libX11")]
    public static extern IntPtr XOpenDisplay(IntPtr display);

    [DllImport("libX11")]
    public static extern IntPtr XDefaultRootWindow(IntPtr display);

    public static IntPtr GetRootWindow() {
        IntPtr display = XOpenDisplay(IntPtr.Zero);
        if (display == IntPtr.Zero) {
            return IntPtr.Zero;
        }
        return XDefaultRootWindow(display);
    }
}
'@

              Add-Type -TypeDefinition $x11Code -Language CSharp

              [IntPtr] $ParentWindow = [X11Interop]::GetRootWindow()

              if ($ParentWindow -ne [System.IntPtr]::Zero) {
                [void] $AquireTokenParameters.WithParentActivityOrWindow($ParentWindow)
              }
            } catch {
              # Do nothing
            }
          }


          #if ($Account) { [void] $AquireTokenParameters.WithAccount($Account) }
          if ($extraScopesToConsent) { [void] $AquireTokenParameters.WithExtraScopesToConsent($extraScopesToConsent) }
          if ($LoginHint) { [void] $AquireTokenParameters.WithLoginHint($LoginHint) }
          if ($Prompt) { [void] $AquireTokenParameters.WithPrompt([Microsoft.Identity.Client.Prompt]::$Prompt) }
          if ($PSBoundParameters.ContainsKey('UseEmbeddedWebView')) { [void] $AquireTokenParameters.WithUseEmbeddedWebView($UseEmbeddedWebView) }

          if (-not $UseEmbeddedWebView) {
            $SystemWebViewOptions = @{}

            if (-not [string]::IsNullOrWhiteSpace($BrowserRedirectSuccess)) {
              $SystemWebViewOptions.BrowserRedirectSuccess = $BrowserRedirectSuccess
            }

            if (-not [string]::IsNullOrWhiteSpace($BrowserRedirectError)) {
              $SystemWebViewOptions.BrowserRedirectError = $BrowserRedirectError
            }

            if (-not [string]::IsNullOrWhiteSpace($HtmlMessageSuccess)) {
              $SystemWebViewOptions.HtmlMessageSuccess = $HtmlMessageSuccess
            }

            if (-not [string]::IsNullOrWhiteSpace($HtmlMessageError)) {
              $SystemWebViewOptions.HtmlMessageError = $HtmlMessageError
            }

            if (@($SystemWebViewOptions.getEnumerator()).count -gt 0) {
              [void] $AquireTokenParameters.WithSystemWebViewOptions($(New-Object Microsoft.Identity.Client.SystemWebViewOptions -Property $SystemWebViewOptions))
            }
          }

          if (!$Timeout -and (($PSBoundParameters.ContainsKey('UseEmbeddedWebView') -and !$UseEmbeddedWebView) -or $PSVersionTable.PSEdition -eq 'Core')) {
            $Timeout = New-TimeSpan -Minutes 2
          }
        } elseif ($PSBoundParameters.ContainsKey('IntegratedWindowsAuth') -and $IntegratedWindowsAuth) {
          $AquireTokenParameters = $PublicClientApplication.AcquireTokenByIntegratedWindowsAuth($Scopes)
          if ($LoginHint) { [void] $AquireTokenParameters.WithUsername($LoginHint) }
        } elseif ($PSBoundParameters.ContainsKey('Silent') -and $Silent) {
          if ($PSBoundParameters.ContainsKey('LoginHint') -and $LoginHint) {
            $AquireTokenParameters = $PublicClientApplication.AcquireTokenSilent($Scopes, $LoginHint)
          } else {
            if ($PSBoundParameters.ContainsKey('AuthenticationBroker') -and $AuthenticationBroker) {
              [Microsoft.Identity.Client.IAccount] $Account = [Microsoft.Identity.Client.PublicClientApplication]::OperatingSystemAccount
            } else {
              [Microsoft.Identity.Client.IAccount] $Account = $PublicClientApplication.GetAccountsAsync().GetAwaiter().GetResult() | Select-Object -First 1
            }

            $AquireTokenParameters = $PublicClientApplication.AcquireTokenSilent($Scopes, $Account)
          }
          if ($PSBoundParameters.ContainsKey('ForceRefresh')) { [void] $AquireTokenParameters.WithForceRefresh($ForceRefresh) }
        } else {
          $paramGetMsalToken = Select-PsBoundParameters -NamedParameter $PSBoundParameters -CommandName 'Get-MsalToken' -CommandParameterSet 'PublicClient-InputObject' -ExcludeParameters 'PublicClientApplication'
          ## Try Silent Authentication
          Write-Verbose ('Attempting Silent Authentication to Application with ClientId [{0}]' -f $ClientApplication.AppConfig.ClientId)
          try {
            $AuthenticationResult = Get-MsalToken -Silent -PublicClientApplication $PublicClientApplication @paramGetMsalToken
            ## Check for requested scopes
            if (CheckForMissingScopes $AuthenticationResult $Scopes) {
              $AuthenticationResult = Get-MsalToken -Interactive -PublicClientApplication $PublicClientApplication @paramGetMsalToken
            }
          } catch [Microsoft.Identity.Client.MsalUiRequiredException] {
            Write-Debug ('{0}: {1}' -f $_.Exception.GetType().Name, $_.Exception.Message)
            ## Try Integrated Windows Authentication
            Write-Verbose ('Attempting Integrated Windows Authentication to Application with ClientId [{0}]' -f $ClientApplication.AppConfig.ClientId)
            try {
              $AuthenticationResult = Get-MsalToken -IntegratedWindowsAuth -PublicClientApplication $PublicClientApplication @paramGetMsalToken
              ## Check for requested scopes
              if (CheckForMissingScopes $AuthenticationResult $Scopes) {
                $AuthenticationResult = Get-MsalToken -Interactive -PublicClientApplication $PublicClientApplication @paramGetMsalToken
              }
            } catch {
              Write-Debug ('{0}: {1}' -f $_.Exception.GetType().Name, $_.Exception.Message)
              ## Revert to Interactive Authentication
              Write-Verbose ('Attempting Interactive Authentication to Application with ClientId [{0}]' -f $ClientApplication.AppConfig.ClientId)
              $AuthenticationResult = Get-MsalToken -Interactive -PublicClientApplication $PublicClientApplication @paramGetMsalToken
            }
          }
          break
        }
      }
      'ConfidentialClient*' {
        if ($PSBoundParameters.ContainsKey('AuthorizationCode')) {
          $AquireTokenParameters = $ConfidentialClientApplication.AcquireTokenByAuthorizationCode($Scopes, $AuthorizationCode)
        } elseif ($PSBoundParameters.ContainsKey('UserAssertion')) {
          if ($UserAssertionType) { [Microsoft.Identity.Client.UserAssertion] $UserAssertionObj = New-Object Microsoft.Identity.Client.UserAssertion -ArgumentList $UserAssertion, $UserAssertionType }
          else { [Microsoft.Identity.Client.UserAssertion] $UserAssertionObj = New-Object Microsoft.Identity.Client.UserAssertion -ArgumentList $UserAssertion }
          $AquireTokenParameters = $ConfidentialClientApplication.AcquireTokenOnBehalfOf($Scopes, $UserAssertionObj)
        } else {
          $AquireTokenParameters = $ConfidentialClientApplication.AcquireTokenForClient($Scopes)
          if ($PSBoundParameters.ContainsKey('ForceRefresh')) { [void] $AquireTokenParameters.WithForceRefresh($ForceRefresh) }
        }
        if ($SendX5C) { [void] $AquireTokenParameters.WithSendX5C($SendX5C) }
      }
      '*' {
        if ($AzureCloudInstance -and $TenantId) { [void] $AquireTokenParameters.WithAuthority($AzureCloudInstance, $TenantId) }
        elseif ($AzureCloudInstance) { [void] $AquireTokenParameters.WithAuthority($AzureCloudInstance, 'common') }
        elseif ($TenantId) { [void] $AquireTokenParameters.WithAuthority($ClientApplication.AppConfig.Authority.AuthorityInfo.CanonicalAuthority.AbsoluteUri, $TenantId) }
        if ($Authority) { [void] $AquireTokenParameters.WithAuthority($Authority.AbsoluteUri) }
        if ($CorrelationId) { [void] $AquireTokenParameters.WithCorrelationId($CorrelationId) }
        if ($ExtraQueryParameters) { [void] $AquireTokenParameters.WithExtraQueryParameters((ConvertTo-Dictionary $ExtraQueryParameters -KeyType ([string]) -ValueType ([string]))) }
        if ($ProofOfPossession) { [void] $AquireTokenParameters.WithProofOfPosession($ProofOfPossession) }
        Write-Debug ('Aquiring Token for Application with ClientId [{0}]' -f $ClientApplication.AppConfig.ClientId)
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
                try { WatchCatchableExitSignal } catch { }

                Start-Sleep -Seconds 1
              } else {
                $tokenSource.Cancel()
                try { $taskAuthenticationResult.Wait() }
                catch { }
                Write-Error -Exception (New-Object System.TimeoutException) -Category ([System.Management.Automation.ErrorCategory]::OperationTimeout) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'GetMsalTokenFailureOperationTimeout' -TargetObject $AquireTokenParameters -ErrorAction Stop
              }
            }
          } finally {
            if (!$taskAuthenticationResult.IsCompleted) {
              Write-Debug ('Canceling Token Acquisition for Application with ClientId [{0}]' -f $ClientApplication.AppConfig.ClientId)
              $tokenSource.Cancel()
            }

            $tokenSource.Dispose()

            if ($InteractiveAuthTopLevelParentWindow) {
              try {
                $InteractiveAuthTopLevelParentWindow.Close()
              } catch {
                # Do nothing
              }
            }
          }

          ## Parse task results
          if ($taskAuthenticationResult.IsFaulted) {
            Write-Error -Exception $taskAuthenticationResult.Exception -Category ([System.Management.Automation.ErrorCategory]::AuthenticationError) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'GetMsalTokenFailureAuthenticationError' -TargetObject $AquireTokenParameters -ErrorAction Stop
          }
          if ($taskAuthenticationResult.IsCanceled) {
            Write-Error -Exception (New-Object System.Threading.Tasks.TaskCanceledException $taskAuthenticationResult) -Category ([System.Management.Automation.ErrorCategory]::OperationStopped) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'GetMsalTokenFailureOperationStopped' -TargetObject $AquireTokenParameters -ErrorAction Stop
          } else {
            $AuthenticationResult = $taskAuthenticationResult.Result
          }
        } catch {
          Write-Error -Exception (Coalesce $_.Exception.InnerException, $_.Exception) -Category ([System.Management.Automation.ErrorCategory]::AuthenticationError) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'GetMsalTokenFailureAuthenticationError' -TargetObject $AquireTokenParameters -ErrorAction Stop
        }
        break
      }
    }

    return $AuthenticationResult
  }
}
