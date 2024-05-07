<#
.SYNOPSIS
    Get client applications from local session cache.
.DESCRIPTION
    This cmdlet will return client applications from the local session cache.
.EXAMPLE
    PS C:\>Get-MsalClientApplication
    Get all client applications in the local session cache.
.EXAMPLE
    PS C:\>Get-MsalClientApplication -ClientId '00000000-0000-0000-0000-000000000000'
    Get client application with specific ClientId from local session cache.
#>
function Get-MsalClientApplication {
    [CmdletBinding(DefaultParameterSetName = 'PublicClient')]
    [OutputType([Microsoft.Identity.Client.PublicClientApplication], [Microsoft.Identity.Client.ConfidentialClientApplication])]
    param
    (
        # Identifier of the client requesting the token.
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $ClientId,
        # Secure secret of the client requesting the token.
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret', ValueFromPipelineByPropertyName = $true)]
        [securestring] $ClientSecret,
        # Client assertion certificate of the client requesting the token.
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientCertificate', ValueFromPipelineByPropertyName = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $ClientCertificate,
        # Address to return to upon receiving a response from the authority.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [uri] $RedirectUri,
        # Tenant identifier of the authority to issue token.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string] $TenantId,
        # Address of the authority to issue token.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [uri] $Authority
    )

    [System.Collections.Generic.List[Microsoft.Identity.Client.IClientApplicationBase]] $listClientApplications = New-Object System.Collections.Generic.List[Microsoft.Identity.Client.IClientApplicationBase]

    switch -Wildcard ($PSCmdlet.ParameterSetName) {
        'PublicClient*' {
            foreach ($PublicClientApplication in $PublicClientApplications) {
                if ((!$ClientId -or $PublicClientApplication.ClientId -eq $ClientId) -and (!$RedirectUri -or $PublicClientApplication.AppConfig.RedirectUri -eq $RedirectUri) -and (!$TenantId -or $PublicClientApplication.AppConfig.TenantId -eq $TenantId) -and (!$Authority -or $PublicClientApplication.Authority -eq $Authority)) {
                    $listClientApplications.Add($PublicClientApplication)
                }
            }

            #$listClientApplications.AddRange(($PublicClientApplications | Where-Object ClientId -eq $ClientId))
        }
        '*' {
            foreach ($ConfidentialClientApplication in $ConfidentialClientApplications) {
                if ((!$ClientId -or $ConfidentialClientApplication.ClientId -eq $ClientId) -and (!$RedirectUri -or $ConfidentialClientApplication.AppConfig.RedirectUri -eq $RedirectUri) -and (!$TenantId -or $ConfidentialClientApplication.AppConfig.TenantId -eq $TenantId) -and (!$Authority -or $ConfidentialClientApplication.Authority -eq $Authority)) {
                    switch ($PSCmdlet.ParameterSetName) {
                        'ConfidentialClientSecret' {
                            if ($ConfidentialClientApplication.AppConfig.ClientSecret -eq $ClientSecret) {
                                $listClientApplications.Add($ConfidentialClientApplication)
                            }
                            break
                        }
                        'ConfidentialClientCertificate' {
                            if ($ConfidentialClientApplication.AppConfig.ClientCredentialCertificate -eq $ClientCertificate) {
                                $listClientApplications.Add($ConfidentialClientApplication)
                            }
                            break
                        }
                        Default {
                            $listClientApplications.Add($ConfidentialClientApplication)
                        }
                    }
                }
            }
        }
    }

    return $listClientApplications
}
