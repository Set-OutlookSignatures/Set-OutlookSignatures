<#
.SYNOPSIS
    Import Configuration
.EXAMPLE
    PS C:\>Import-Config
    Import Configuration
.INPUTS
    System.String
#>
function Import-Config {
    [CmdletBinding()]
    [OutputType([psobject])]
    param (
        # Configuration File Path
        [Parameter(Mandatory = $false)]
        [string] $Path = 'config.json'
    )

    ## Initialize
    #if (![IO.Path]::IsPathRooted($Path)) {
    #    $AppDataDirectory = Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ApplicationData)) 'MSAL.PS'
    #    $Path = Join-Path $AppDataDirectory $Path
    #}

    if (Test-Path -LiteralPath $Path) {
        ## Load from File
        $ModuleConfigPersistent = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json

        ## Return Config
        return $ModuleConfigPersistent
    }
}
