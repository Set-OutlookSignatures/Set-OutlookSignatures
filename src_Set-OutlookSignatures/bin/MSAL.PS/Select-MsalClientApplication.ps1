<#
.SYNOPSIS
    Select the closest match from local session cache based on the client application parameters provided and create one if there is no match.
.DESCRIPTION
    This cmdlet should only be used if you want to store your client application in the local session cache.
.EXAMPLE
    PS C:\>Select-MsalClientApplication -ClientId '00000000-0000-0000-0000-000000000000'
    Select a public client application from local session cache with client id 00000000-0000-0000-0000-000000000000 and default settings.
#>
function Select-MsalClientApplication {
    [CmdletBinding(DefaultParameterSetName = 'PublicClient')]
    [OutputType([Microsoft.Identity.Client.PublicClientApplication], [Microsoft.Identity.Client.ConfidentialClientApplication])]
    param
    (
        # Identifier of the client requesting the token.
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject', Position = 1, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject', Position = 1, ValueFromPipelineByPropertyName = $true)]
        [string] $ClientId,
        # Secure secret of the client requesting the token.
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        #[AllowNull()]
        [securestring] $ClientSecret,
        # Client assertion certificate of the client requesting the token.
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $ClientCertificate,
        # Address to return to upon receiving a response from the authority.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [uri] $RedirectUri,
        # Instance of Azure Cloud
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.Identity.Client.AzureCloudInstance] $AzureCloudInstance,
        # Tenant identifier of the authority to issue token.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string] $TenantId,
        # Address of the authority to issue token.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [uri] $Authority,
        # Use Platform Authentication Broker
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [switch] $AuthenticationBroker,
        # Public client application options
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-InputObject', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.Identity.Client.PublicClientApplicationOptions] $PublicClientOptions,
        # Confidential client application options
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClient-InputObject', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.Identity.Client.ConfidentialClientApplicationOptions] $ConfidentialClientOptions
    )

    $paramNewMsalClientApplication = Select-PsBoundParameters $PSBoundParameters -CommandName New-MsalClientApplication -ExcludeParameters ErrorAction
    $NewClientApplication = New-MsalClientApplication -ErrorAction Stop @paramNewMsalClientApplication
    switch -Wildcard ($PSCmdlet.ParameterSetName) {
        "PublicClient*" {
            #[Microsoft.Identity.Client.IPublicClientApplication] $ClientApplication = $PublicClientApplications | Where-Object { $_.ClientId -eq $NewClientApplication.ClientId -and $_.AppConfig.RedirectUri -eq $NewClientApplication.AppConfig.RedirectUri -and $_.AppConfig.TenantId -eq $NewClientApplication.AppConfig.TenantId } | Select-Object -Last 1
            [Microsoft.Identity.Client.IPublicClientApplication] $ClientApplication = $PublicClientApplications | Where-Object { $_.ClientId -eq $NewClientApplication.ClientId -and $_.AppConfig.RedirectUri -eq $NewClientApplication.AppConfig.RedirectUri -and $_.AppConfig.IsBrokerEnabled -eq $NewClientApplication.AppConfig.IsBrokerEnabled } | Select-Object -Last 1
            break
        }
        "ConfidentialClientSecret" {
            #[Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplications | Where-Object { $_.ClientId -eq $NewClientApplication.ClientId -and $_.AppConfig.ClientSecret -eq $NewClientApplication.AppConfig.ClientSecret -and $_.AppConfig.RedirectUri -eq $NewClientApplication.AppConfig.RedirectUri -and $_.AppConfig.TenantId -eq $NewClientApplication.AppConfig.TenantId } | Select-Object -Last 1
            [Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplications | Where-Object { $_.ClientId -eq $NewClientApplication.ClientId -and $_.AppConfig.ClientSecret -eq $NewClientApplication.AppConfig.ClientSecret -and $_.AppConfig.RedirectUri -eq $NewClientApplication.AppConfig.RedirectUri } | Select-Object -Last 1
            break
        }
        "ConfidentialClientCertificate" {
            #[Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplications | Where-Object { $_.ClientId -eq $NewClientApplication.ClientId -and $_.AppConfig.ClientCredentialCertificate -eq $NewClientApplication.AppConfig.ClientCredentialCertificate -and $_.AppConfig.RedirectUri -eq $NewClientApplication.AppConfig.RedirectUri -and $_.AppConfig.TenantId -eq $NewClientApplication.AppConfig.TenantId } | Select-Object -Last 1
            [Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplications | Where-Object { $_.ClientId -eq $NewClientApplication.ClientId -and $_.AppConfig.ClientCredentialCertificate -eq $NewClientApplication.AppConfig.ClientCredentialCertificate -and $_.AppConfig.RedirectUri -eq $NewClientApplication.AppConfig.RedirectUri } | Select-Object -Last 1
            break
        }
        "ConfidentialClient-InputObject" {
            if ($NewClientApplication.AppConfig.ClientSecret) {
                #[Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplications | Where-Object { $_.ClientId -eq $NewClientApplication.ClientId -and $_.AppConfig.ClientSecret -eq $NewClientApplication.AppConfig.ClientSecret -and $_.AppConfig.RedirectUri -eq $NewClientApplication.AppConfig.RedirectUri -and $_.AppConfig.TenantId -eq $NewClientApplication.AppConfig.TenantId } | Select-Object -Last 1
                [Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplications | Where-Object { $_.ClientId -eq $NewClientApplication.ClientId -and $_.AppConfig.ClientSecret -eq $NewClientApplication.AppConfig.ClientSecret -and $_.AppConfig.RedirectUri -eq $NewClientApplication.AppConfig.RedirectUri } | Select-Object -Last 1
            }
            else {
                #[Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplications | Where-Object { $_.ClientId -eq $NewClientApplication.ClientId -and $_.AppConfig.ClientCredentialCertificate -eq $NewClientApplication.AppConfig.ClientCredentialCertificate -and $_.AppConfig.RedirectUri -eq $NewClientApplication.AppConfig.RedirectUri -and $_.AppConfig.TenantId -eq $NewClientApplication.AppConfig.TenantId } | Select-Object -Last 1
                [Microsoft.Identity.Client.IConfidentialClientApplication] $ClientApplication = $ConfidentialClientApplications | Where-Object { $_.ClientId -eq $NewClientApplication.ClientId -and $_.AppConfig.ClientCredentialCertificate -eq $NewClientApplication.AppConfig.ClientCredentialCertificate -and $_.AppConfig.RedirectUri -eq $NewClientApplication.AppConfig.RedirectUri } | Select-Object -Last 1
            }
            break
        }
    }

    if (!$ClientApplication) {
        $ClientApplication = $NewClientApplication
        Write-Debug ('Adding Application with ClientId [{0}] and RedirectUri [{1}] to cache.' -f $ClientApplication.AppConfig.ClientId, $ClientApplication.AppConfig.RedirectUri)
        Add-MsalClientApplication $ClientApplication
    }
    else {
        Write-Debug ('Using Application with ClientId [{0}] and RedirectUri [{1}] from cache.' -f $ClientApplication.AppConfig.ClientId, $ClientApplication.AppConfig.RedirectUri)
    }

    return $ClientApplication
}
