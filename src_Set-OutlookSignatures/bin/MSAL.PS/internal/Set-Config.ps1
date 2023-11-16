<#
.SYNOPSIS
    Set Configuration
.EXAMPLE
    PS C:\>Set-Config
    Set Configuration
.INPUTS
    System.String
#>
function Set-Config {
    [CmdletBinding()]
    #[OutputType([psobject])]
    param (
        # Configuration Object
        [Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $true)]
        [psobject] $InputObject,
        # Allow use of previously loaded dlls
        [Parameter(Mandatory = $false)]
        [bool] $DllLenientLoading,
        # Prompt user when previously loaded dlls conflict
        [Parameter(Mandatory = $false)]
        [bool] $DllLenientLoadingPrompt,
        # Read settings from environment variables
        [Parameter(Mandatory = $false)]
        [switch] $ResolveFromEnvironmentVariables,
        # Variable to output config
        [Parameter(Mandatory = $false)]
        [ref] $OutConfig = ([ref]$script:ModuleConfig)
    )

    ## Update local configuration
    if ($ResolveFromEnvironmentVariables) {
        if (${env:msalps.dll.lenientLoading}) {
            $OutConfig.Value.'dll.lenientLoading' = ${env:msalps.dll.lenientLoading}
            $OutConfig.Value.'dll.lenientLoadingPrompt' = $false
        }
        if (${env:msalps.dll.lenientLoadingPrompt}) { $OutConfig.Value.'dll.lenientLoadingPrompt' = ${env:msalps.dll.lenientLoadingPrompt} }
    }
    if ($InputObject) {
        if ($InputObject -is [hashtable]) { $InputObject = [PSCustomObject]$InputObject }
        foreach ($Property in $InputObject.psobject.Properties) {
            if ($OutConfig.Value.psobject.Properties.Name -contains $Property.Name) {
                $OutConfig.Value.($Property.Name) = $Property.Value
            }
            else {
                Write-Warning ('Ignoring invalid configuration property [{0}].' -f $Property.Name)
            }
        }
    }
    if ($PSBoundParameters.ContainsKey('DllLenientLoading')) { $OutConfig.Value.'dll.lenientLoading' = $DllLenientLoading }
    if ($PSBoundParameters.ContainsKey('DllLenientLoadingPrompt')) { $OutConfig.Value.'dll.lenientLoadingPrompt' = $DllLenientLoadingPrompt }

    ## Return updated local configuration
    #return $OutConfig.Value
}
