<#
.SYNOPSIS
    Add client application to local session cache.
.DESCRIPTION
    This cmdlet will add a client application object to the local session cache.
.EXAMPLE
    PS C:\>Add-MsalClientApplication $ClientApplication
    Add client application to the local session cache.
.EXAMPLE
    PS C:\>Add-MsalClientApplication $ClientApplication -PassThru
    Add client application to the local session cache and return application object.
#>
function Add-MsalClientApplication {
    [CmdletBinding(DefaultParameterSetName = 'PublicClient')]
    [OutputType([Microsoft.Identity.Client.PublicClientApplication], [Microsoft.Identity.Client.ConfidentialClientApplication])]
    param
    (
        # Public client application
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient', Position = 0, ValueFromPipeline = $true)]
        [Microsoft.Identity.Client.IPublicClientApplication] $PublicClientApplication,
        # Confidential client application
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClient', Position = 0, ValueFromPipeline = $true)]
        [Microsoft.Identity.Client.IConfidentialClientApplication] $ConfidentialClientApplication,
        # Returns client application
        [Parameter(Mandatory = $false)]
        [switch] $PassThru
    )

    switch ($PSCmdlet.ParameterSetName) {
        "PublicClient" {
            $ClientApplication = $PublicClientApplication
            if (!$PublicClientApplications.Contains($ClientApplication)) {
                $PublicClientApplications.Add($ClientApplication)
            }
            else {
                $Exception = New-Object ArgumentException -ArgumentList 'The client application provided already exists in the session cache.'
                Write-Error -Exception $Exception -Category ([System.Management.Automation.ErrorCategory]::ResourceExists) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'AddMsalClientApplicationFailureAlreadyExists' -TargetObject $ClientApplication #-ErrorAction Stop
            }
            break
        }
        "ConfidentialClient" {
            $ClientApplication = $ConfidentialClientApplication
            if (!$ConfidentialClientApplications.Contains($ClientApplication)) {
                $ConfidentialClientApplications.Add($ClientApplication)
            }
            else {
                $Exception = New-Object ArgumentException -ArgumentList 'The client application provided already exists in the session cache.'
                Write-Error -Exception $Exception -Category ([System.Management.Automation.ErrorCategory]::ResourceExists) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'AddMsalClientApplicationFailureAlreadyExists' -TargetObject $ClientApplication #-ErrorAction Stop
            }
            break
        }
    }

    if ($PassThru) {
        Write-Output $ClientApplication
    }
}
