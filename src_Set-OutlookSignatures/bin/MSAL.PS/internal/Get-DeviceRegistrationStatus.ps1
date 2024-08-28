<#
.SYNOPSIS
    Get Entra ID Device Registration Status from current device
.EXAMPLE
    PS C:\>Get-DeviceRegistrationStatus
    Get Entra ID Device Registration Status from current device
.INPUTS
    System.String
#>
function Get-DeviceRegistrationStatus {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param ()

    ## Get Device Registration Status
    [hashtable] $Dsreg = @{}

    if (([System.Environment]::OSVersion.Platform -eq 'Win32NT') -and ([System.Environment]::OSVersion.Version -ge '10.0') -and ([System.Environment]::OSVersion.Version.Build -ge 17134)) {
        try {
            dsregcmd /status | ForEach-Object { if ($_ -match '\s*(.+) : (.+)') { $Dsreg.Add($Matches[1], $Matches[2]) } }
        } catch {
        }
    }

    return $Dsreg
}
