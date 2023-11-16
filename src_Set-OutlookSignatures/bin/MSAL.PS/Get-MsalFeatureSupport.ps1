<#
.SYNOPSIS
    List supported features on current platform and session.
.DESCRIPTION

.EXAMPLE
    PS C:\>Get-MsalFeatureSupport
    List supported features on current platform and session.
#>
function Get-MsalFeatureSupport {
    [CmdletBinding()]
    param ()

    return [PSCustomObject]$script:ModuleFeatureSupport
}
