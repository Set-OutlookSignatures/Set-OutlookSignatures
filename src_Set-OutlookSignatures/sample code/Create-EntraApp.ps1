<#
This sample code shows how to automate the creation of the Entra ID app required for Set-OutlookSignatures.

Both types of apps are supported: The one for end users, and the one for SimulateAndDeploy.

You can adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers professional support for this and other open source code.
#>


#Requires -Version 5.1

[CmdletBinding()]

param (
    # Which type of app should be created?
    #   'Set-OutlookSignatures' for the default Set-OutlookSignatures app being accessed by end users runnding Set-OutlookSignatures
    #     Uses only delegated permissions, as described in '.\config\default graph config.ps1'
    #   'SimulateAndDeploy' for use in the "simulate and deploy" scenario
    #     Uses delegated permissions and application permissions, as described in '.\sample code\SimulateAndDeploy.ps1'
    #   For security reasons, the app type has no default value and needs to be set manually
    [ValidateSet('Set-OutlookSignatures', 'SimulateAndDeploy', 'OutlookAddIn')]
    $AppType = $null,

    [ValidateNotNullOrEmpty()]
    $AppName = $null,

    [ValidateNotNullOrEmpty()]
    [uri]$OutlookAddInUrl = $null
)


Clear-Host

# Remove unnecessary ETS type data associated with arrays in Windows PowerShell
Remove-TypeData System.Array -ErrorAction SilentlyContinue

if ($psISE) {
    Write-Host 'PowerShell ISE detected. Use PowerShell in console or terminal instead.' -ForegroundColor Red
    Write-Host 'Required features are not available in ISE. Exit.' -ForegroundColor Red
    exit 1
}

if (($ExecutionContext.SessionState.LanguageMode) -ine 'FullLanguage') {
    Write-Host "This PowerShell session runs in $($ExecutionContext.SessionState.LanguageMode) mode, not FullLanguage mode." -ForegroundColor Red
    Write-Host 'Required features are only available in FullLanguage mode. Exit.' -ForegroundColor Red
    exit 1
}

$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

if ($AppName) {
    $AppName = $AppName.trim()
}

Write-Host 'Set-OutlookSignatures Create-EntraApp.ps1'

$ParameterCheckSuccess = $true

if ([string]::IsNullOrWhiteSpace($AppType)) {
    $ParameterCheckSuccess = $false

    Write-Host '  App type not defined, exiting.' -ForegroundColor Red
    Write-Host "    Add parameter '-AppType' with one of the following values: $(($PSCmdlet.MyInvocation.MyCommand.Parameters['AppType'].Attributes |
    Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }).ValidValues -join ', ') " -ForegroundColor Red
}

if ([string]::IsNullOrWhiteSpace($AppName)) {
    $ParameterCheckSuccess = $false

    Write-Host '  App name not defined, exiting.' -ForegroundColor Red
    Write-Host "    Add parameter '-AppName' with a name for the Entra ID app to be created." -ForegroundColor Red
}

if ($AppType -ieq 'OutlookAddIn' -and $OutlookAddInUrl -eq $null) {
    $ParameterCheckSuccess = $false

    Write-Host '  Outlook Add-In URI not defined, exiting.' -ForegroundColor Red
    Write-Host "    Add parameter '-OutlookAddInUrl' with a URI for the Outlook add-in to be created." -ForegroundColor Red
}

if (($AppType -iin @('Set-OutlookSignatures', 'SimulateAndDeploy')) -and ($OutlookAddInUrl -ne $null)) {
    Write-Host "  Outlook Add-In URI not allowed for app type $($AppType), exiting." -ForegroundColor Red
    Write-Host "    Remove parameter '-OutlookAddInUrl'." -ForegroundColor Red
    exit 1
}

if (-not $ParameterCheckSuccess) {
    exit 1
} else {
    $AppName = $AppName.trim()
}


Write-Host
Write-Host 'Create Entra ID app'
Write-Host "  App type: $($AppType)"
Write-Host "  App name: $($AppName)"
if ($AppType -ieq 'OutlookAddIn') {
    Write-Host "  Outlook Add-In URI: $($OutlookAddInUrl)"
}


Write-Host
Write-Host 'Install Microsoft.Graph PowerShell modules'
foreach ($MicrosoftGraphPowerShellModule in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications', 'Microsoft.Graph.Identity.SignIns')) {
    Write-Host "  $($MicrosoftGraphPowerShellModule)"

    if (Get-Module -ListAvailable -Name $MicrosoftGraphPowerShellModule) {
        Update-Module $MicrosoftGraphPowerShellModule
    } else {
        Install-Module $MicrosoftGraphPowerShellModule -Scope CurrentUser -Force -AllowClobber
    }
}


Write-Host
Write-Host "Connect to your Entra ID with a user being 'Application Adminstrator' or 'Global Administrator'"
Write-Host '  An authentication window will open, likely in a browser'

# Disconnect first, so that no existing connection is re-used. This forces to choose an account for the following connect.
$null = Disconnect-MgGraph -ErrorAction SilentlyContinue

Connect-MgGraph -ContextScope Process -Scopes 'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All' -NoWelcome


Write-Host
Write-Host 'Create a new app registration'
Write-Host "  App name: $($AppName)"

$ExistingApp = @(Get-MgApplication | Where-Object { $_.DisplayName.Trim() -ieq $AppName })

if ($ExistingApp.Count -gt 0) {
    $ExistingApp | ForEach-Object {
        Write-Host "  App with name '$($AppName)' already exists. ID: $($_.Id)" -ForegroundColor Red
    }

    Write-Host '  Exiting.' -ForegroundColor Red
    exit 1
}

$params = @{
    DisplayName    = $AppName
    Description    = "$($AppType) app for Set-OutlookSignatures: Email signatures and out-of-office replies for Exchange and Outlook. Full-featured, cost-effective, unsurpassed data privacy."
    Notes          = "$($AppType) app for Set-OutlookSignatures: Email signatures and out-of-office replies for Exchange and Outlook. Full-featured, cost-effective, unsurpassed data privacy."
    SignInAudience = 'AzureADMyOrg'
}

$app = New-MgApplication @params

Write-Host
Write-Host 'Add required permissions to app registration'
if ($AppType -ieq 'Set-OutlookSignatures') {
    $permissionParams = @{
        RequiredResourceAccess = @(
            @{
                # Microsoft Graph
                'ResourceAppId'  = '00000003-0000-0000-c000-000000000000'
                'ResourceAccess' = @(
                    # Microsoft Graph permissions reference: https://learn.microsoft.com/en-us/graph/permissions-reference

                    # Delegated permission: email
                    #   Allows the app to read your users' primary email address.
                    #   Required to log on the current user.
                    @{
                        'id'   = '64a6cdd6-aab1-4aaf-94b8-3cc8405e90d0'
                        'type' = 'Scope'
                    },

                    # Delegated permission: EWS.AccessAsUser.All
                    #   Allows the app to have the same access to mailboxes as the signed-in user via Exchange Web Services.
                    #   Required to connect to Outlook Web and to set Outlook Web signature (classic and roaming).
                    @{
                        'id'   = '9769c687-087d-48ac-9cb3-c37dde652038'
                        'type' = 'Scope'
                    },

                    # Delegated permission: Files.Read.All
                    #   Allows the app to read all files the signed-in user can access.
                    #   Required for access to templates and configuration files hosted on SharePoint Online.
                    #   For added security, use Files.SelectedOperations.Selected as alternative, requiring granting specific permissions in SharePoint Online.
                    @{
                        'id'   = 'df85f4d6-205c-4ac5-a5ea-6bf408dba283'
                        'type' = 'Scope'
                    },

                    # Delegated permission: GroupMember.Read.All
                    #   Allows the app to list groups, read basic group properties and read membership of all groups the signed-in user has access to.
                    #   Required to find groups by name and to get their security identifier (SID) and the number of transitive members.
                    @{
                        'id'   = 'bc024368-1153-4739-b217-4326f2e966d0'
                        'type' = 'Scope'
                    },

                    # Delegated permission: Mail.ReadWrite
                    #   Allows the app to create, read, update, and delete email in user mailboxes. Does not include permission to send mail.
                    #   Required to connect to Outlook Web and to set Outlook signatures.
                    @{
                        'id'   = '024d486e-b451-40bb-833d-3e66d98c5c73'
                        'type' = 'Scope'
                    },

                    # Delegated permission: MailboxSettings.ReadWrite
                    #   Allows the app to create, read, update, and delete user's mailbox settings. Does not include permission to send mail.
                    #   Required to detect the state of the out-of-office assistant and to set out-of-office replies.
                    @{
                        'id'   = '818c620a-27a9-40bd-a6a5-d96f7d610b4b'
                        'type' = 'Scope'
                    },

                    # Delegated permission: offline_access
                    #   Allows the app to see and update the data you gave it access to, even when users are not currently using the app. This does not give the app any additional permissions.
                    #   Required to get a refresh token from Graph.
                    @{
                        'id'   = '7427e0e9-2fba-42fe-b0c0-848c9e6a8182'
                        'type' = 'Scope'
                    },

                    # Delegated permission: openid
                    #   Allows users to sign in to the app with their work or school accounts and allows the app to see basic user profile information.
                    #   Required to log on the current user.
                    @{
                        'id'   = '37f7f235-527c-4136-accd-4a02d197296e'
                        'type' = 'Scope'
                    },

                    # Delegated permission: profile
                    #   Allows the app to see your users' basic profile (e.g., name, picture, user name, email address).
                    #   Required to log on the current user, to access the '/me' Graph API, to get basic properties of the current user.
                    @{
                        'id'   = '14dad69e-099b-42c9-810b-d002981feec1'
                        'type' = 'Scope'
                    },

                    # Delegated permission: User.Read.All
                    #   Allows the app to read the full set of profile properties, reports, and managers of other users in your organization, on behalf of the signed-in user.
                    #   Required for $CurrentUser[…]$ and $CurrentMailbox[…]$ replacement variables, and for simulation mode.
                    @{
                        'id'   = 'a154be20-db9c-4678-8ab7-66f6cc099a59'
                        'type' = 'Scope'
                    }
                )
            }
        )
    }
} elseif ($AppType -ieq 'SimulateAndDeploy') {
    $permissionParams = @{
        RequiredResourceAccess = @(
            @{
                # Microsoft Graph
                'resourceAppId'  = '00000003-0000-0000-c000-000000000000'
                'resourceAccess' = @(
                    # Microsoft Graph permissions reference: https://learn.microsoft.com/en-us/graph/permissions-reference

                    # Microsoft Graph permissions reference: https://learn.microsoft.com/en-us/graph/permissions-reference

                    # Delegated permission: email
                    #   Allows the app to read your users' primary email address.
                    #   Required to log on the current user.
                    @{
                        'id'   = '64a6cdd6-aab1-4aaf-94b8-3cc8405e90d0'
                        'type' = 'Scope'
                    },

                    # Delegated permission: EWS.AccessAsUser.All
                    #   Allows the app to have the same access to mailboxes as the signed-in user via Exchange Web Services.
                    #   Required to connect to Outlook Web and to set Outlook Web signature (classic and roaming).
                    @{
                        'id'   = '9769c687-087d-48ac-9cb3-c37dde652038'
                        'type' = 'Scope'
                    },

                    # Delegated permission: Files.Read.All
                    #   Allows the app to read all files the signed-in user can access.
                    #   Required for access to SharePoint Online on Linux, macOS, and on Windows without WebDAV.
                    #   You can use Files.SelectedOperations.Selected as alternative, requiring granting specific permission in SharePoint Online.
                    @{
                        'id'   = 'df85f4d6-205c-4ac5-a5ea-6bf408dba283'
                        'type' = 'Scope'
                    },

                    # Delegated permission: GroupMember.Read.All
                    #   Allows the app to list groups, read basic group properties and read membership of all groups the signed-in user has access to.
                    #   Required to find groups by name and to get their security identifier (SID) and the number of transitive members.
                    @{
                        'id'   = 'bc024368-1153-4739-b217-4326f2e966d0'
                        'type' = 'Scope'
                    },

                    # Delegated permission: Mail.ReadWrite
                    #   Allows the app to create, read, update, and delete email in user mailboxes. Does not include permission to send mail.
                    #   Required to connect to Outlook Web and to set Outlook signatures.
                    @{
                        'id'   = '024d486e-b451-40bb-833d-3e66d98c5c73'
                        'type' = 'Scope'
                    },

                    # Delegated permission: MailboxSettings.ReadWrite
                    #   Allows the app to create, read, update, and delete user's mailbox settings. Does not include permission to send mail.
                    #   Required to detect the state of the out-of-office assistant and to set out-of-office replies.
                    @{
                        'id'   = '818c620a-27a9-40bd-a6a5-d96f7d610b4b'
                        'type' = 'Scope'
                    },

                    # Delegated permission: offline_access
                    #   Allows the app to see and update the data you gave it access to, even when users are not currently using the app. This does not give the app any additional permissions.
                    #   Required to get a refresh token from Graph.
                    @{
                        'id'   = '7427e0e9-2fba-42fe-b0c0-848c9e6a8182'
                        'type' = 'Scope'
                    },

                    # Delegated permission: openid
                    #   Allows users to sign in to the app with their work or school accounts and allows the app to see basic user profile information.
                    #   Required to log on the current user.
                    @{
                        'id'   = '37f7f235-527c-4136-accd-4a02d197296e'
                        'type' = 'Scope'
                    },

                    # Delegated permission: profile
                    #   Allows the app to see your users' basic profile (e.g., name, picture, user name, email address).
                    #   Required to log on the current user, to access the '/me' Graph API, to get basic properties of the current user.
                    @{
                        'id'   = '14dad69e-099b-42c9-810b-d002981feec1'
                        'type' = 'Scope'
                    },

                    # Delegated permission: User.Read.All
                    #   Allows the app to read the full set of profile properties, reports, and managers of other users in your organization, on behalf of the signed-in user.
                    #   Required for $CurrentUser[…]$ and $CurrentMailbox[…]$ replacement variables, and for simulation mode.
                    @{
                        'id'   = 'a154be20-db9c-4678-8ab7-66f6cc099a59'
                        'type' = 'Scope'
                    },

                    # Application permission: Files.Read.All
                    #   Allows the app to read all files in all site collections without a signed in user.
                    #   Required for access to templates and configuration files hosted on SharePoint Online.
                    #   For added security, use Files.SelectedOperations.Selected as alternative, requiring granting specific permissions in SharePoint Online.
                    @{
                        'id'   = '01d4889c-1287-42c6-ac1f-5d1e02578ef6'
                        'type' = 'Role'
                    },

                    # Application permission: GroupMember.Read.All
                    #   Allows the app to list groups, read basic group properties and read membership of all groups the signed-in user has access to.
                    #   Required to find groups by name and to get their security identifier (SID) and the number of transitive members.
                    @{
                        'id'   = '98830695-27a2-44f7-8c18-0c3ebc9698f6'
                        'type' = 'Role'
                    },

                    # Application permission: Mail.ReadWrite
                    #   Allows the app to create, read, update, and delete mail in all mailboxes without a signed-in user. Does not include permission to send mail.
                    #   Required to connect to Outlook Web and to set Outlook signatures.
                    @{
                        'id'   = 'e2a3a72e-5f79-4c64-b1b1-878b674786c9'
                        'type' = 'Role'
                    },

                    # Application permission: MailboxSettings.ReadWrite
                    #   Allows the app to create, read, update, and delete user's mailbox settings. Does not include permission to send mail.
                    #   Required to detect the state of the out-of-office assistant and to set out-of-office replies.
                    @{
                        'id'   = '6931bccd-447a-43d1-b442-00a195474933'
                        'type' = 'Role'
                    },

                    # Application permission: User.Read.All
                    #   Allows the app to read the full set of profile properties, reports, and managers of other users in your organization, on behalf of the signed-in user.
                    #   Required for $CurrentUser[…]$ and $CurrentMailbox[…]$ replacement variables, and for simulation mode.
                    @{
                        'id'   = 'df021288-bdef-4463-88db-98f22de89214'
                        'type' = 'Role'
                    }
                )
            },
            @{
                # Office 365 Exchange Online
                'resourceAppId'  = '00000002-0000-0ff1-ce00-000000000000'
                'resourceAccess' = @(
                    @{
                        # Application permission: full_access_as_app
                        #   Allows the app to have full access via Exchange Web Services to all mailboxes without a signed-in user.
                        #   Required for Exchange Web Services access (read Outlook Web configuration, set classic signature and roaming signatures)
                        'id'   = 'dc890d15-9560-4a4c-9b7f-a736ec74ec40'
                        'type' = 'Role'
                    }
                )
            }
        )
    }
} elseif ($AppType -ieq 'OutlookAddin') {
    $permissionParams = @{
        RequiredResourceAccess = @(
            @{
                # Microsoft Graph
                'resourceAppId'  = '00000003-0000-0000-c000-000000000000'
                'resourceAccess' = @(
                    # Microsoft Graph permissions reference: https://learn.microsoft.com/en-us/graph/permissions-reference

                    # Delegated permission: GroupMember.Read.All
                    #   Allows the app to list groups, read basic group properties and read membership of all groups the signed-in user has access to.
                    #   Required to find and check license groups.
                    @{
                        'id'   = 'bc024368-1153-4739-b217-4326f2e966d0'
                        'type' = 'Scope'
                    },

                    # Delegated permission: Mail.Read
                    #   Allows to read emails in mailbox of the currently logged-on user (and in no other mailboxes).
                    #    Required because of Microsoft restrictions accessing roaming signatures.
                    @{
                        'id'   = '570282fd-fa5c-430d-a7fd-fc8dc98a9dca'
                        'type' = 'Scope'
                    },

                    # Delegated permission: User.Read.All
                    #   Allows the app to read the full set of profile properties, reports, and managers of other users in your organization, on behalf of the signed-in user.
                    #   Required to get the UPN for a given SMTP email address.
                    @{
                        'id'   = 'a154be20-db9c-4678-8ab7-66f6cc099a59'
                        'type' = 'Scope'
                    }
                )
            }
        )
    }
}

Update-MgApplication -ApplicationId $app.Id -BodyParameter $permissionParams

if ($AppType -iin @('Set-OutlookSignatures', 'SimulateAndDeploy')) {
    Write-Host '  Consider restricting file access by switching from Files.Read.All to Files.SelectedOperations.Selected.'
    Write-Host '    This enhances security but requires granting specific permissions in SharePoint Online.'
}


Write-Host
Write-Host 'Add redirect URIs to app registration'
if ($AppType -iin @('Set-OutlookSignatures', 'SimulateAndDeploy')) {
    $params =	@{
        RedirectUris = @(
            'http://localhost',
            "ms-appx-web://microsoft.aad.brokerplugin/$($app.AppId)"
        )
    }

    Update-MgApplication -ApplicationId $app.Id -PublicClient $params
} elseif ($AppType -ieq 'OutlookAddIn') {
    $params =	@{
        RedirectUris = @(
            "brk-multihub://$($OutlookAddInUrl.DnsSafeHost)"
        )
    }

    Update-MgApplication -ApplicationId $app.Id -Spa $params
}

$params.RedirectUris | ForEach-Object {
    Write-Host "  $($_)"
}

if ($AppType -iin @('Set-OutlookSignatures', 'SimulateAndDeploy')) {
    Write-Host
    Write-Host 'Enable public client flow'

    Update-MgApplication -ApplicationId $app.Id -IsFallbackPublicClient
}


if ($AppType -ieq 'SimulateAndDeploy') {
    Write-Host
    Write-Host 'Add client secret to app registration'

    $params = @{
        displayName = "Initial client secret, valid $(Get-Date -Format 'yyyy-MM-dd')--$(Get-Date (Get-Date).AddMonths(24) -Format 'yyyy-MM-dd')"
        endDateTime = (Get-Date).AddMonths(24)
    }

    $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $params
}


Write-Host
Write-Host 'Grant admin consent'
Write-Host ' This may take a moment'
$AppServicePrincipal = New-MgServicePrincipal -AppId $App.AppId
$delegatedPermissions = @{}

foreach ($resource in $permissionParams.RequiredResourceAccess) {
    foreach ($resourcePermission in $resource.resourceAccess) {
        if ($resourcePermission.type -eq 'Role') {
            # Application permission

            $null = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AppServicePrincipal.Id -PrincipalId $AppServicePrincipal.Id -ResourceId $((Get-MgServicePrincipal -Filter "AppId eq '$($resource.resourceAppId)'").Id) -AppRoleId $resourcePermission.id
        } elseif ($resourcePermission.type -eq 'Scope') {
            # Delegated permission

            $delegatedPermissions[$((Get-MgServicePrincipal -Filter "AppId eq '$($resource.resourceAppId)'").Id)] += " $(((Get-MgServicePrincipal -Filter "appId eq '$($resource.resourceAppId)'").Oauth2PermissionScopes | Where-Object { $_.Id -eq $resourcePermission.id }).Value)"
        }
    }
}

$delegatedPermissions.GetEnumerator() | ForEach-Object {
    $null = New-MgOauth2PermissionGrant -ClientId $AppServicePrincipal.Id -ConsentType 'AllPrincipals' -ResourceId $_.Key -Scope $_.Value.trim()
}


Write-Host
Write-Host 'Disconnect from Entra ID'
$null = Disconnect-MgGraph -ErrorAction SilentlyContinue


Write-Host
Write-Host '▼▼▼ Relevant information for your configuration below ▼▼▼' -ForegroundColor Green
if ($AppType -ieq 'Set-OutlookSignatures') {
    Write-Host "  GraphClientId for Set-OutlookSignatures: '$($app.AppId)'"
} elseif ($AppType -ieq 'SimulateAndDeploy') {
    Write-Host "  GraphClientId for SimulateAndDeploy: '$($app.AppId)'"
    Write-Host "  GraphClientSecret for SimulateAndDeploy: '$($secret.SecretText)'"
    Write-Host "    Do not forget to renew the client secret before $(Get-Date (Get-Date).AddMonths(24) -Format 'yyyy-MM-dd')!"
} elseif ($AppType -ieq 'OutlookAddIn') {
    Write-Host "  GRAPH_CLIENT_ID for Outlook Add-In: '$($app.AppId)'"
}
Write-Host '▲▲▲ Relevant information for your configuration above ▲▲▲' -ForegroundColor Green

Write-Host
Write-Host 'Done'