<#
.SYNOPSIS
    Get user from token cache of application.
.DESCRIPTION

.EXAMPLE
    PS C:\>$ClientApplication = Get-MsalClientApplication -ClientId '00000000-0000-0000-0000-000000000000'
    PS C:\>$ClientApplication | Get-MsalAccount
    Get all accounts from client application cache.
#>
function Get-MsalAccount {
    [CmdletBinding()]
    param
    (
        # Client application
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'ClientApplication', Position = 0)]
        [Microsoft.Identity.Client.IClientApplicationBase] $ClientApplication,
        # Information of a single account.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Accounts', Position = 0)]
        [Microsoft.Identity.Client.IAccount[]] $Accounts,
        # The username in UserPrincipalName (UPN) format.
        [Parameter(Mandatory = $false)]
        [string] $Username
    )

    if ($PSCmdlet.ParameterSetName -eq 'ClientApplication') {
        [Microsoft.Identity.Client.IAccount[]] $Accounts = $ClientApplication.GetAccountsAsync().GetAwaiter().GetResult()
    }

    if ($Username) {
        return $Accounts | Where-Object Username -EQ $Username
    } else {
        return $Accounts
    }
}
