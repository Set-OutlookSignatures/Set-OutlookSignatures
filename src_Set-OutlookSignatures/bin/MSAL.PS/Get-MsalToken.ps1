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

    # BrowserRedirectError
    [uri] $BrowserRedirectError,

    # BrowserRedirectSuccess
    [uri] $BrowserRedirectSuccess,

    # HtmlMessageSuccess
    [string] $HtmlMessageSuccess,

    # HtmlMessageError
    [string] $HtmlMessageError,

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
        } elseif ($PSBoundParameters.ContainsKey('DeviceCode') -and $DeviceCode -or ($Interactive -and !$script:ModuleFeatureSupport.WebView1Support -and !$script:ModuleFeatureSupport.WebView2Support -and $PublicClientApplication.AppConfig.RedirectUri -ne 'http://localhost/') -or ($Interactive -and !$script:ModuleFeatureSupport.WebView1Support -and $PublicClientApplication.AppConfig.RedirectUri -eq 'urn:ietf:wg:oauth:2.0:oob')) {
          $AquireTokenParameters = $PublicClientApplication.AcquireTokenWithDeviceCode($Scopes, [DeviceCodeHelper]::GetDeviceCodeResultCallback())
        } elseif ($PSBoundParameters.ContainsKey('Interactive') -and $Interactive) {
          $AquireTokenParameters = $PublicClientApplication.AcquireTokenInteractive($Scopes)
          [IntPtr] $ParentWindow = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle
          if ($ParentWindow -eq [System.IntPtr]::Zero -and [System.Environment]::OSVersion.Platform -eq 'Win32NT') {
            # Detect parent window even when run in Windows Terminal
            # https://github.com/german-one/termwnd
            # Copyright (c) Steffen Illhardt
            # Licensed under the MIT license.

            # min. req.: PowerShell v.2

            try {
              Add-Type -EA SilentlyContinue -TypeDefinition @'
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

//# provides properties identifying the terminal window the current console application is running in
public static class WinTerm {
  //# imports the used Windows API functions
  private static class NativeMethods {
    [DllImport("kernel32.dll")]
    internal static extern int CloseHandle(IntPtr Hndl);
    [DllImport("kernelbase.dll")]
    internal static extern int CompareObjectHandles(IntPtr hFirst, IntPtr hSecond);
    [DllImport("kernel32.dll")]
    internal static extern int DuplicateHandle(IntPtr SrcProcHndl, IntPtr SrcHndl, IntPtr TrgtProcHndl, out IntPtr TrgtHndl, int Acc, int Inherit, int Opts);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool EnumWindows(EnumWindowsProc enumFunc, IntPtr lparam);
    [DllImport("user32.dll")]
    internal static extern IntPtr GetAncestor(IntPtr hWnd, int flgs);
    [DllImport("kernel32.dll")]
    internal static extern IntPtr GetConsoleWindow();
    [DllImport("kernel32.dll")]
    internal static extern IntPtr GetCurrentProcess();
    [DllImport("user32.dll")]
    internal static extern IntPtr GetWindow(IntPtr hWnd, int cmd);
    [DllImport("user32.dll")]
    internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint procId);
    [DllImport("ntdll.dll")]
    internal static extern int NtQuerySystemInformation(int SysInfClass, IntPtr SysInf, int SysInfLen, out int RetLen);
    [DllImport("kernel32.dll")]
    internal static extern IntPtr OpenProcess(int Acc, int Inherit, uint ProcId);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    internal static extern int QueryFullProcessImageNameW(IntPtr Proc, int Flgs, StringBuilder Name, ref int Size);
    [DllImport("user32.dll")]
    internal static extern IntPtr SendMessageW(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
  }

  private static readonly IntPtr conWnd = NativeMethods.GetConsoleWindow();

  public static IntPtr HWnd { get { return hWnd; } } //# window handle
  public static uint Pid { get { return pid; } } //# process id
  public static uint Tid { get { return tid; } } //# thread id
  public static string BaseName { get { return baseName; } } //# process name without .exe extension

  private static IntPtr hWnd = IntPtr.Zero;
  private static uint pid = 0;
  private static uint tid = 0;
  private static string baseName = string.Empty;

  //# owns an unmanaged resource
  //# the ctor qualifies a SafeRes object to manage either a pointer received from Marshal.AllocHGlobal(), or a handle
  private class SafeRes : CriticalFinalizerObject, IDisposable {
    //# resource type of a SafeRes object
    internal enum ResType { MemoryPointer, Handle }

    private IntPtr raw = IntPtr.Zero;
    private readonly ResType resourceType = ResType.MemoryPointer;

    internal IntPtr Raw { get { return raw; } }
    internal bool IsInvalid { get { return raw == IntPtr.Zero || raw == new IntPtr(-1); } }

    //# constructs a SafeRes object from an unmanaged resource specified by parameter raw
    //# the resource must be either a pointer received from Marshal.AllocHGlobal() (specify resourceType ResType.MemoryPointer),
    //# or a handle (specify resourceType ResType.Handle)
    internal SafeRes(IntPtr raw, ResType resourceType) {
      this.raw = raw;
      this.resourceType = resourceType;
    }

    ~SafeRes() { Dispose(false); }

    public void Dispose() {
      Dispose(true);
      GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing) {
      if (IsInvalid) { return; }
      if (resourceType == ResType.MemoryPointer) {
        Marshal.FreeHGlobal(raw);
        raw = IntPtr.Zero;
        return;
      }

      if (NativeMethods.CloseHandle(raw) != 0) { raw = new IntPtr(-1); }
    }

    internal virtual void Reset(IntPtr raw) {
      Dispose();
      this.raw = raw;
    }
  }

  //# undocumented SYSTEM_HANDLE structure, SYSTEM_HANDLE_TABLE_ENTRY_INFO might be the actual name
  [StructLayout(LayoutKind.Sequential)]
  private struct SystemHandle {
    internal readonly uint ProcId; //# PID of the process the SYSTEM_HANDLE belongs to
    internal readonly byte ObjTypeId; //# identifier of the object
    internal readonly byte Flgs;
    internal readonly ushort Handle; //# value representing an opened handle in the process
    internal readonly IntPtr pObj;
    internal readonly uint Acc;
  }

  private static string GetProcBaseName(SafeRes sHProc) {
    int size = 1024;
    StringBuilder nameBuf = new StringBuilder(size);
    return NativeMethods.QueryFullProcessImageNameW(sHProc.Raw, 0, nameBuf, ref size) == 0 ? "" : Path.GetFileNameWithoutExtension(nameBuf.ToString(0, size));
  }

  //# Enumerate the opened handles in each process, select those that refer to the same process as findOpenProcId.
  //# Return the ID of the process that opened the handle if its name is the same as searchProcName,
  //# Return 0 if no such process is found.
  private static uint GetPidOfNamedProcWithOpenProcHandle(string searchProcName, uint findOpenProcId) {
    const int PROCESS_DUP_HANDLE = 0x0040, //# access right to duplicate handles
              PROCESS_QUERY_LIMITED_INFORMATION = 0x1000, //# access right to retrieve certain process information
              STATUS_INFO_LENGTH_MISMATCH = unchecked((int)0xc0000004), //# NTSTATUS returned if we still didn't allocate enough memory
              SystemHandleInformation = 16; //# one of the SYSTEM_INFORMATION_CLASS values
    const byte OB_TYPE_INDEX_JOB = 7; //# one of the SYSTEM_HANDLE.ObjTypeId values
    int status, //# retrieves the NTSTATUS return value
        infSize = 0x200000, //# initially allocated memory size for the SYSTEM_HANDLE_INFORMATION object
        len;

    //# allocate some memory representing an undocumented SYSTEM_HANDLE_INFORMATION object, which can't be meaningfully declared in C# code
    using (SafeRes sPSysHndlInf = new SafeRes(Marshal.AllocHGlobal(infSize), SafeRes.ResType.MemoryPointer)) {
      //# try to get an array of all available SYSTEM_HANDLE objects, allocate more memory if necessary
      while ((status = NativeMethods.NtQuerySystemInformation(SystemHandleInformation, sPSysHndlInf.Raw, infSize, out len)) == STATUS_INFO_LENGTH_MISMATCH) {
        sPSysHndlInf.Reset(Marshal.AllocHGlobal(infSize = len + 0x1000));
      }

      if (status < 0) { return 0; }
      using (SafeRes sHFindOpenProc = new SafeRes(NativeMethods.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, findOpenProcId), SafeRes.ResType.Handle)) { //# intentionally after NtQuerySystemInformation() was called to exclude it from the found open handles
        if (sHFindOpenProc.IsInvalid) { return 0; }
        uint foundPid = 0, curPid = 0;
        IntPtr hThis = NativeMethods.GetCurrentProcess();
        int sysHndlSize = Marshal.SizeOf(typeof(SystemHandle));
        using (SafeRes sHCur = new SafeRes(IntPtr.Zero, SafeRes.ResType.Handle)) {
          //# iterate over the array of SYSTEM_HANDLE objects, which begins at an offset of pointer size in the SYSTEM_HANDLE_INFORMATION object
          //# the number of SYSTEM_HANDLE objects is specified in the first 32 bits of the SYSTEM_HANDLE_INFORMATION object
          for (IntPtr pSysHndl = (IntPtr)((long)sPSysHndlInf.Raw + IntPtr.Size), pEnd = (IntPtr)((long)pSysHndl + Marshal.ReadInt32(sPSysHndlInf.Raw) * sysHndlSize);
               pSysHndl != pEnd;
               pSysHndl = (IntPtr)((long)pSysHndl + sysHndlSize)) {
            //# get one SYSTEM_HANDLE at a time
            SystemHandle sysHndl = (SystemHandle)Marshal.PtrToStructure(pSysHndl, typeof(SystemHandle));
            //# shortcut; OB_TYPE_INDEX_JOB is the identifier we are looking for, any other SYSTEM_HANDLE object is immediately ignored at this point
            if (sysHndl.ObjTypeId != OB_TYPE_INDEX_JOB) { continue; }
            //# every time the process changes, the previous handle needs to be closed and we open a new handle to the current process
            if (curPid != sysHndl.ProcId) {
              curPid = sysHndl.ProcId;
              sHCur.Reset(NativeMethods.OpenProcess(PROCESS_DUP_HANDLE | PROCESS_QUERY_LIMITED_INFORMATION, 0, curPid));
            }

            //# if the process has not been opened, or
            //# if duplicating the current one of its open handles fails, continue with the next SYSTEM_HANDLE object
            //# the duplicated handle is necessary to get information about the object (e.g. the process) it points to
            IntPtr hCurOpenDup;
            if (sHCur.IsInvalid ||
                NativeMethods.DuplicateHandle(sHCur.Raw, (IntPtr)sysHndl.Handle, hThis, out hCurOpenDup, PROCESS_QUERY_LIMITED_INFORMATION, 0, 0) == 0) {
              continue;
            }

            using (SafeRes sHCurOpenDup = new SafeRes(hCurOpenDup, SafeRes.ResType.Handle)) {
              if (NativeMethods.CompareObjectHandles(sHCurOpenDup.Raw, sHFindOpenProc.Raw) != 0 && //# both the handle of the open process and the currently duplicated handle must refer to the same kernel object
                  searchProcName == GetProcBaseName(sHCur)) { //# the process name of the currently found process must meet the process name we are looking for
                foundPid = curPid;
                break;
              }
            }
          }
        }
        return foundPid;
      }
    }
  }

  private static uint findPid;
  private static IntPtr foundHWnd;

  private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

  private static bool GetOpenConWndCallback(IntPtr hWnd, IntPtr lParam) {
    uint thisPid;
    uint thisTid = NativeMethods.GetWindowThreadProcessId(hWnd, out thisPid);
    if (thisTid == 0 || thisPid != findPid)
      return true;

    foundHWnd = hWnd;
    return false;
  }

  private static IntPtr GetOpenConWnd(uint termPid) {
    if (termPid == 0)
      return IntPtr.Zero;

    findPid = termPid;
    foundHWnd = IntPtr.Zero;
    NativeMethods.EnumWindows(new EnumWindowsProc(GetOpenConWndCallback), IntPtr.Zero);
    return foundHWnd;
  }

  private static IntPtr GetTermWnd(ref bool terminalExpected) {
    const int WM_GETICON = 0x007F,
              GW_OWNER = 4,
              GA_ROOTOWNER = 3;

    //# We don't have a proper way to figure out to what terminal app the Shell process
    //# is connected on the local machine:
    //# https://github.com/microsoft/terminal/issues/7434
    //# We're getting around this assuming we don't get an icon handle from the
    //# invisible Conhost window when the Shell is connected to Windows Terminal.
    terminalExpected = NativeMethods.SendMessageW(conWnd, WM_GETICON, IntPtr.Zero, IntPtr.Zero) == IntPtr.Zero;
    if (!terminalExpected)
      return conWnd;

    //# Polling because it may take some milliseconds for Terminal to create its window and take ownership of the hidden ConPTY window.
    IntPtr conOwner = IntPtr.Zero; //# FWIW this receives the terminal window our tab is created in, but it gets never updated if the tab is moved to another window.
    for (int i = 0; i < 200 && conOwner == IntPtr.Zero; ++i) {
      Thread.Sleep(5);
      conOwner = NativeMethods.GetWindow(conWnd, GW_OWNER);
    }

    //# Something went wrong if polling did not succeed within 1 second (e.g. it's not Windows Terminal).
    if (conOwner == IntPtr.Zero)
      return IntPtr.Zero;

    //# Get the ID of the Shell process that spawned the Conhost process.
    uint shellPid;
    uint shellTid = NativeMethods.GetWindowThreadProcessId(conWnd, out shellPid);
    if (shellTid == 0)
      return IntPtr.Zero;

    //# Get the ID of the OpenConsole process spawned for the Shell process.
    uint openConPid = GetPidOfNamedProcWithOpenProcHandle("OpenConsole", shellPid);
    if (openConPid == 0)
      return IntPtr.Zero;

    //# Get the hidden window of the OpenConsole process
    IntPtr openConWnd = GetOpenConWnd(openConPid);
    if (openConWnd == IntPtr.Zero)
      return IntPtr.Zero;

    //# The root owner window is the Terminal window.
    return NativeMethods.GetAncestor(openConWnd, GA_ROOTOWNER);
  }

  static WinTerm() {
    Refresh();
  }

  //# used to initially get or to update the properties if a terminal tab is moved to another window
  public static void Refresh() {
    const int PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;
    bool terminalExpected = false;
    hWnd = GetTermWnd(ref terminalExpected);
    if (hWnd == IntPtr.Zero)
      throw new InvalidOperationException();

    tid = NativeMethods.GetWindowThreadProcessId(hWnd, out pid);
    if (tid == 0)
      throw new InvalidOperationException();

    using (SafeRes sHProc = new SafeRes(NativeMethods.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, pid), SafeRes.ResType.Handle)) {
      if (sHProc.IsInvalid)
        throw new InvalidOperationException();

      baseName = GetProcBaseName(sHProc);
    }

    if (string.IsNullOrEmpty(baseName) || (terminalExpected && baseName != "WindowsTerminal"))
      throw new InvalidOperationException();
  }
}
'@
            } catch {}

            if ('WinTerm' -as [type]) {
              [IntPtr] $ParentWindow = [int] "0X$([WinTerm]::HWnd.ToString('X8'))"
            } else {
              $ParentWindow = [System.IntPtr]::Zero
            }
          }
          if ($ParentWindow -ne [System.IntPtr]::Zero) { [void] $AquireTokenParameters.WithParentActivityOrWindow($ParentWindow) }
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
          if ($LoginHint) {
            $AquireTokenParameters = $PublicClientApplication.AcquireTokenSilent($Scopes, $LoginHint)
          } else {
            [Microsoft.Identity.Client.IAccount] $Account = $PublicClientApplication.GetAccountsAsync().GetAwaiter().GetResult() | Select-Object -First 1
            $AquireTokenParameters = $PublicClientApplication.AcquireTokenSilent($Scopes, $Account)
          }
          if ($PSBoundParameters.ContainsKey('ForceRefresh')) { [void] $AquireTokenParameters.WithForceRefresh($ForceRefresh) }
        } else {
          $paramGetMsalToken = Select-PsBoundParameters -NamedParameter $PSBoundParameters -CommandName 'Get-MsalToken' -CommandParameterSet 'PublicClient-InputObject' -ExcludeParameters 'PublicClientApplication'
          ## Try Silent Authentication
          Write-Verbose ('Attempting Silent Authentication to Application with ClientId [{0}]' -f $ClientApplication.ClientId)
          try {
            $AuthenticationResult = Get-MsalToken -Silent -PublicClientApplication $PublicClientApplication @paramGetMsalToken
            ## Check for requested scopes
            if (CheckForMissingScopes $AuthenticationResult $Scopes) {
              $AuthenticationResult = Get-MsalToken -Interactive -PublicClientApplication $PublicClientApplication @paramGetMsalToken
            }
          } catch [Microsoft.Identity.Client.MsalUiRequiredException] {
            Write-Debug ('{0}: {1}' -f $_.Exception.GetType().Name, $_.Exception.Message)
            ## Try Integrated Windows Authentication
            Write-Verbose ('Attempting Integrated Windows Authentication to Application with ClientId [{0}]' -f $ClientApplication.ClientId)
            try {
              $AuthenticationResult = Get-MsalToken -IntegratedWindowsAuth -PublicClientApplication $PublicClientApplication @paramGetMsalToken
              ## Check for requested scopes
              if (CheckForMissingScopes $AuthenticationResult $Scopes) {
                $AuthenticationResult = Get-MsalToken -Interactive -PublicClientApplication $PublicClientApplication @paramGetMsalToken
              }
            } catch {
              Write-Debug ('{0}: {1}' -f $_.Exception.GetType().Name, $_.Exception.Message)
              ## Revert to Interactive Authentication
              Write-Verbose ('Attempting Interactive Authentication to Application with ClientId [{0}]' -f $ClientApplication.ClientId)
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
              } else {
                $tokenSource.Cancel()
                try { $taskAuthenticationResult.Wait() }
                catch { }
                Write-Error -Exception (New-Object System.TimeoutException) -Category ([System.Management.Automation.ErrorCategory]::OperationTimeout) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'GetMsalTokenFailureOperationTimeout' -TargetObject $AquireTokenParameters -ErrorAction Stop
              }
            }
          } finally {
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
