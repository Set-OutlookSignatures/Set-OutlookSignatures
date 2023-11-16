<#
.SYNOPSIS
    Convert/Decrypt SecureString to Plain Text String.
.EXAMPLE
    PS C:\>ConvertFrom-SecureStringAsPlainText (ConvertTo-SecureString 'SuperSecretString' -AsPlainText -Force) -Force
    Convert plain text to SecureString and then convert it back.
.INPUTS
    System.Security.SecureString
#>
function ConvertFrom-SecureStringAsPlainText {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        # Secure String Value
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [securestring] $SecureString,
        # Confirms that you understand the implications of using the AsPlainText parameter and still want to use it.
        [Parameter(Mandatory = $true)]
        [switch] $Force
    )

    try {
        [IntPtr] $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        Write-Output ([System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR))
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    }
}
