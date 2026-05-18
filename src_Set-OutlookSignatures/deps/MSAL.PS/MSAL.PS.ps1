$BridgeCode = @'
using System;
using System.Reflection;
using System.IO;

namespace MSALPS.Loader {
    public static class DependencyResolver {
        private static string _basePath;
        private static ResolveEventHandler _handler;

        public static void Initialize(string path) {
            _basePath = path;
            _handler = new ResolveEventHandler(Handler);
            AppDomain.CurrentDomain.AssemblyResolve += _handler;
        }

        public static void Uninitialize() {
            if (_handler != null) {
                AppDomain.CurrentDomain.AssemblyResolve -= _handler;
                _handler = null;
            }
        }

        private static Assembly Handler(object sender, ResolveEventArgs args) {
            AssemblyName name = new AssemblyName(args.Name);

            if (name.Name == "System.Memory") {
                string dll = Path.Combine(_basePath, "System.Memory.dll");
                return File.Exists(dll) ? Assembly.LoadFrom(dll) : null;
            }

            string pathCheck = Path.Combine(_basePath, name.Name + ".dll");
            if (File.Exists(pathCheck)) {
                return Assembly.LoadFrom(pathCheck);
            }
            return null;
        }
    }
}
'@

if ($PSVersionTable.PSVersion.Major -le 5) {
    if (-not ([System.Type]::GetType('MSALPS.Loader.DependencyResolver'))) {
        Add-Type -TypeDefinition $BridgeCode
    }

    [MSALPS.Loader.DependencyResolver]::Initialize((Join-Path -Path $PSScriptRoot -ChildPath 'netstandard2.0'))
}


#region Import Helper Functions
function Catch-AssemblyLoadError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $AssemblyPath
    )

    ## Save ErrorRecord to throw later
    $ErrorRecord = $_

    ## Look for existing assembly
    [string] $AssemblyName = [System.IO.Path]::GetFileName($AssemblyPath)
    $Assembly = [System.AppDomain]::CurrentDomain.GetAssemblies().Where{ $AssemblyName -eq $_.ManifestModule }
    if (-not $Assembly) { throw $ErrorRecord }

    ## On older Windows OSes, the desktop DLL conflicts with itself so just ignore.
    if ($Assembly.Location.StartsWith($PSScriptRoot)) { return }

    Write-Warning (@'
Assembly with same name "{0}" is already loaded:
{1}
'@ -f $AssemblyName, $Assembly.Location)

    ## Ask the user
    if ($script:ModuleConfig.'dll.lenientLoadingPrompt') {
        $DefaultChoice = if ($script:ModuleConfig.'dll.lenientLoading') { 1 } else { 2 }
        $DllLenientLoading = Write-HostPrompt 'Ignore assembly conflict and continue importing module?' -Message 'Some module functionality will not work.' -Choices @('&Yes', '&No') -DefaultChoice $DefaultChoice -ErrorAction SilentlyContinue
        if ($DllLenientLoading -eq 1) {
            $script:ModuleConfig.'dll.lenientLoading' = $true

            $PersistModuleConfig = Write-HostPrompt 'Remember settings?' -Message ('Module settings will be persisted in "{0}"' -f (Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ApplicationData)) '/MSAL.PS/config.json')) -DefaultChoice 2 -Choices @('&Yes', '&No') -ErrorAction SilentlyContinue
            if ($PersistModuleConfig -eq 1) {
                $script:ModuleConfig.'dll.lenientLoadingPrompt' = $false
                Export-Config
            } else {
                Write-Host @'

# You may also suppress this prompt by providing module settings on import:
Import-Module MSAL.PS -ArgumentList @{ 'dll.lenientLoading' = $true; 'dll.lenientLoadingPrompt' = $false }

# Or defining the following environment variable:
${env:msalps.dll.lenientLoading} = $true # Continue Module Import

'@
            }
        } else { $script:ModuleConfig.'dll.lenientLoading' = $false }
    }

    ## Throw error if strict dll loading
    if (!$script:ModuleConfig.'dll.lenientLoading') { throw $ErrorRecord }
    else { $script:ModuleFeatureSupport.WebView2Support = $false }

    return $Assembly.Location
}

#endregion Import Helper Functions

## Read Module Manifest
$ModuleManifest = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'MSAL.PS.psd1')
[System.Collections.Generic.List[string]] $RequiredAssemblies = New-Object System.Collections.Generic.List[string]


## Select the correct assemblies for the PowerShell platform
foreach ($Path in @($ModuleManifest.FileList -ilike '*\netstandard2.0\Microsoft.Identity.Client.dll')) {
    $RequiredAssemblies.Add($([System.IO.Path]::GetFullPath($(Join-Path -Path $PSScriptRoot -ChildPath $Path))))
}

if ($PSVersionTable.PSEdition -eq 'Core') {
    foreach ($Path in @($ModuleManifest.FileList -ilike '*\netstandard2.0\Microsoft.Identity*.dll')) {
        $RequiredAssemblies.Add($([System.IO.Path]::GetFullPath($(Join-Path -Path $PSScriptRoot -ChildPath $Path))))
    }
} elseif ($PSVersionTable.PSEdition -eq 'Desktop') {
    foreach ($Path in @($ModuleManifest.FileList -ilike '*\netstandard2.0\Microsoft.Identity*.dll')) {
        $RequiredAssemblies.Add($([System.IO.Path]::GetFullPath($(Join-Path -Path $PSScriptRoot -ChildPath $Path))))
    }
}

foreach ($RequiredAssembly in $RequiredAssemblies) {
    try {
        Add-Type -LiteralPath $RequiredAssembly -IgnoreWarnings | Out-Null
    } catch {
        $RequiredAssembly = Catch-AssemblyLoadError $RequiredAssembly
    }
}


# Load TokenCacheHelper
if ([System.Environment]::OSVersion.Platform -eq 'Win32NT') {
    if (-not ('TokenCacheHelper' -as [type])) {
        foreach ($Path in ($ModuleManifest.FileList -like '*\internal\TokenCacheHelper.cs')) {
            $srcTokenCacheHelper = [System.IO.Path]::GetFullPath($(Join-Path -Path $PSScriptRoot -ChildPath $Path))
        }

        if ($PSVersionTable.PSVersion -ge [version]'7.0') {
            $RequiredAssemblies.AddRange([string[]]('netstandard.dll', 'System.Threading.dll', 'System.Runtime.Extensions.dll', 'System.IO.FileSystem.dll', 'System.Security.Cryptography.ProtectedData.dll'))
            Add-Type -LiteralPath $srcTokenCacheHelper -ReferencedAssemblies $RequiredAssemblies
        } elseif ($PSVersionTable.PSVersion -ge [version]'5.1') {
            $RequiredAssemblies.AddRange([string[]]('netstandard.dll', 'System.Security.dll'))
            Add-Type -LiteralPath $srcTokenCacheHelper -ReferencedAssemblies $RequiredAssemblies
        }
    }
}


# Load DeviceCodeHelper
if (-not ('DeviceCodeHelper' -as [type])) {
    foreach ($Path in ($ModuleManifest.FileList -like '*\internal\DeviceCodeHelper.cs')) {
        $srcDeviceCodeHelper = [System.IO.Path]::GetFullPath($(Join-Path -Path $PSScriptRoot -ChildPath $Path))
    }
    if ($PSVersionTable.PSVersion -ge [version]'6.0') {
        $RequiredAssemblies.Add('System.Console.dll')
    }

    try {
        Add-Type -LiteralPath $srcDeviceCodeHelper -ReferencedAssemblies $RequiredAssemblies -IgnoreWarnings -WarningAction SilentlyContinue
    } catch {
        $script:ModuleFeatureSupport.DeviceCodeSupport = $false
        Write-Warning 'There was an error loading some dependencies. DeviceCode parameter will not function.'
    }
}


<#
if ($PSVersionTable.PSVersion.Major -le 5) {
    # This removes the event handler from the AppDomain to prevent
    # memory leaks or conflicts with other modules later in the session.
    [MSALPS.Loader.DependencyResolver]::Uninitialize()
}
#>