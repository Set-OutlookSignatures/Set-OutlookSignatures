<#
.SYNOPSIS
    Remove client application from local session cache.
.DESCRIPTION
    This cmdlet will remove a client application object from the local session cache.
.EXAMPLE
    PS C:\>Remove-MsalClientApplication $ClientApplication
    Remove specified client application from local session cache.
#>
function Remove-MsalClientApplication {
    [CmdletBinding(DefaultParameterSetName = 'PublicClient')]
    [OutputType([Microsoft.Identity.Client.PublicClientApplication], [Microsoft.Identity.Client.ConfidentialClientApplication])]
    param
    (
        # Public client application
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient', Position = 0, ValueFromPipeline = $true)]
        [Microsoft.Identity.Client.IPublicClientApplication] $PublicClientApplication,
        # Confidential client application
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClient', Position = 0, ValueFromPipeline = $true)]
        [Microsoft.Identity.Client.IConfidentialClientApplication] $ConfidentialClientApplication
    )

    switch ($PSCmdlet.ParameterSetName) {
        'PublicClient' {
            $ClientApplication = $PublicClientApplication
            $Result = $PublicClientApplications.Remove($ClientApplication)
            break
        }
        'ConfidentialClient' {
            $ClientApplication = $ConfidentialClientApplication
            $Result = $ConfidentialClientApplications.Remove($ClientApplication)
            break
        }
    }

    if (!$Result) {
        $Exception = New-Object ArgumentException -ArgumentList 'The client application provided was not found in session cache.'
        Write-Error -Exception $Exception -Category ([System.Management.Automation.ErrorCategory]::ObjectNotFound) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'RemoveMsalClientApplicationFailureNotFound' -TargetObject $ClientApplication -ErrorAction Stop
    }
}
