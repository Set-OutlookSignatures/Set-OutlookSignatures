<#
.SYNOPSIS
    Convert hashtable to generic dictionary.
.DESCRIPTION

.EXAMPLE
    PS C:\>ConvertTo-Dictionary @{ KeyName = 'StringValue' } -ValueType ([string])
    Convert hashtable to generic dictionary.
.INPUTS
    System.Hashtable
#>
function ConvertTo-Dictionary {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[object, object]])]
    param (
        # Value to convert
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [hashtable[]] $InputObjects,
        # Data Type of Key
        [Parameter(Mandatory = $false)]
        [type] $KeyType = [string],
        # Data Type of Value
        [Parameter(Mandatory = $false)]
        [type] $ValueType = [object]
    )

    process {
        foreach ($InputObject in $InputObjects) {
            $OutputObject = New-Object ('System.Collections.Generic.Dictionary[[{0}],[{1}]]' -f $KeyType.FullName, $ValueType.FullName)
            foreach ($KeyPair in $InputObject.GetEnumerator()) {
                $OutputObject.Add($KeyPair.Key, $KeyPair.Value)
            }

            Write-Output $OutputObject
        }
    }
}
