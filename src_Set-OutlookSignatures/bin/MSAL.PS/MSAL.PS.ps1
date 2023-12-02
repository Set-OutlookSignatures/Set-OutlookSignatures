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
$ModuleManifest = Import-PowerShellDataFile (Join-Path $PSScriptRoot 'MSAL.PS.psd1')
[System.Collections.Generic.List[string]] $RequiredAssemblies = New-Object System.Collections.Generic.List[string]

## Select the correct assemblies for the PowerShell platform
# Having .net5 and netcoreapp dlls causes an import error when they are both listed in the filelist.
# if ($PSVersionTable.PSVersion -ge [version]'7.1' -and $IsWindows -and $PSVersionTable.OS -match '\d+(\.\d+)+$' -and [version]$matches[0] -ge [version]'10.0.17763') {
#     foreach ($Path in ($ModuleManifest.FileList -like "*\Microsoft.Identity.Client.*\net5.0-windows10.0.17763\*.dll")) {
#         $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
#     }
# }
#if ($PSVersionTable.PSEdition -eq 'Core') {
    foreach ($Path in ($ModuleManifest.FileList -ilike '*\netstandard2.0\*.dll')) {
        $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
    }
#} elseif ($PSVersionTable.PSEdition -eq 'Desktop') {
#    foreach ($Path in ($ModuleManifest.FileList -ilike '*\netframework\*.dll')) {
#        $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
#    }
#}

## Load correct assemblies for the PowerShell platform
foreach ($RequiredAssembly in $RequiredAssemblies) {
    try {
        Add-Type -LiteralPath $RequiredAssembly -IgnoreWarnings | Out-Null
    } catch {
        $RequiredAssembly = Catch-AssemblyLoadError $RequiredAssembly
    }
}


## Load TokenCacheHelper
if ([System.Environment]::OSVersion.Platform -eq 'Win32NT') {
    foreach ($Path in ($ModuleManifest.FileList -like '*\internal\TokenCacheHelper.cs')) {
        $srcTokenCacheHelper = Join-Path $PSScriptRoot $Path
    }

    if ($PSVersionTable.PSEdition -ieq 'Core') {
        $RequiredAssemblies.AddRange([string[]]@('netstandard.dll', 'System.Threading.dll', 'System.Runtime.Extensions.dll', 'System.IO.FileSystem.dll', 'System.Security.Cryptography.ProtectedData.dll', 'mscorlib.dll'))

        if (-not ('TokenCacheHelper' -as [type])) {
            try {
                Add-Type -LiteralPath $srcTokenCacheHelper -ReferencedAssemblies $RequiredAssemblies
            } catch {
                Write-Warning 'There was an error loading some dependencies. Storing TokenCache on disk will not function.'
            }
        }
    } else {
        $RequiredAssemblies.AddRange([string[]]@('netstandard.dll', 'System.Security.dll'))

        if (-not ('TokenCacheHelper' -as [type])) {
            try {
                Add-Type -LiteralPath $srcTokenCacheHelper -ReferencedAssemblies $RequiredAssemblies -IgnoreWarnings
            } catch {
                Write-Warning 'There was an error loading some dependencies. Storing TokenCache on disk will not function.'
            }
        }
    }
}

## Load DeviceCodeHelper
foreach ($Path in ($ModuleManifest.FileList -like '*\internal\DeviceCodeHelper.cs')) {
    $srcDeviceCodeHelper = Join-Path $PSScriptRoot $Path
}
if ($PSVersionTable.PSVersion -ge [version]'6.0') {
    $RequiredAssemblies.Add('System.Console.dll')
    #$RequiredAssemblies.Add('System.ComponentModel.Primitives.dll')
    #$RequiredAssemblies.Add('System.Diagnostics.Process.dll')
}

if (-not ('DeviceCodeHelper' -as [type])) {
    try {
        Add-Type -LiteralPath $srcDeviceCodeHelper -ReferencedAssemblies $RequiredAssemblies -IgnoreWarnings -WarningAction SilentlyContinue
    } catch {
        $script:ModuleFeatureSupport.DeviceCodeSupport = $false
        Write-Warning 'There was an error loading some dependencies. DeviceCode parameter will not function.'
    }
}
