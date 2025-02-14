<#
This sample code shows how to automate the creation of the Entra ID app required for Set-OutlookSignatures.

Both types of apps are supported: The one for end users, and the one for SimulateAndDeploy.

You can adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers commercial support for this and other open source code.
#>

[CmdletBinding()]

param (
    # Which type of app should be created?
    #   'Set-OutlookSignatures' for the default Set-OutlookSignatures app being accessed by end users runnding Set-OutlookSignatures
    #     Uses only delegated permissions, as described in '.\config\default graph config.ps1'
    #   'SimulateAndDeploy' for use in the "simulate and deploy" scenario
    #     Uses delegated permissions and application permissions, as described in '.\sample code\SimulateAndDeploy.ps1'
    #   For security reasons, the app type has no default value and needs to be set manually
    [ValidateSet('Set-OutlookSignatures', 'SimulateAndDeploy')]
    $AppType = $null,

    [ValidateNotNullOrEmpty()]
    $AppName = $null

)


Clear-Host

if ($psISE) {
    Write-Host 'PowerShell ISE detected. Use PowerShell in console or terminal instead.' -ForegroundColor Red
    Write-Host 'Required features are not available in ISE. Exit.' -ForegroundColor Red
    exit 1
}

$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

Set-Location $PSScriptRoot

if ($AppName) {
    $AppName = $AppName.trim()
}

Write-Host 'Create Entra ID app for Set-OutlookSignatures'
Write-Host "  App type: $($AppType)"
Write-Host "  App name: $($AppName)"

if ([string]::IsNullOrWhiteSpace($AppType)) {
    Write-Host
    Write-Host '  App type not defined, exiting.' -ForegroundColor Red
    Write-Host "  Add parameter '-AppType' with one of the following values: $(($PSCmdlet.MyInvocation.MyCommand.Parameters['AppType'].Attributes |
    Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }).ValidValues -join ', ') " -ForegroundColor Red
}

if ([string]::IsNullOrWhiteSpace($AppName)) {
    Write-Host
    Write-Host '  App name not defined, exiting.' -ForegroundColor Red
    Write-Host "  Add parameter '-AppName' with a name for the Entra ID app to be created." -ForegroundColor Red
}

if ([string]::IsNullOrWhiteSpace($AppType) -or [string]::IsNullOrWhiteSpace($AppName)) {
    exit 1
}


Write-Host
Write-Host 'Install Microsoft.Graph PowerShell modules'
foreach ($MicrosoftGraphPowerShellModule in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')) {
    if (Get-Module -ListAvailable -Name $MicrosoftGraphPowerShellModule) {
        Update-Module $MicrosoftGraphPowerShellModule -Scope CurrentUser
    } else {
        Install-Module $MicrosoftGraphPowerShellModule -Scope CurrentUser -Force -AllowClobber
    }
}


Write-Host
Write-Host "Connect to your Entra ID with a user being 'Application Adminstrator' or 'Global Administrator'"
# Disconnect first, so that no existing connection is re-used. This forces to choose an account for the following connect.
$null = Disconnect-MgGraph -ErrorAction SilentlyContinue
Connect-MgGraph -Scopes 'Application.ReadWrite.All' -NoWelcome


Write-Host
Write-Host 'Create a new app registration'
Write-Host '  Does not check if an app with the same name already exists'
Write-Host "  App name: $($AppName)"
$params = @{
    DisplayName    = $AppName
    Description    = 'Set-OutlookSignatures, email signatures and out-of-office replies for Exchange and all of Outlook: Classic and New, Windows, Web, Mac, Linux, Android, iOS'
    Notes          = 'Set-OutlookSignatures, email signatures and out-of-office replies for Exchange and all of Outlook: Classic and New, Windows, Web, Mac, Linux, Android, iOS'
    SignInAudience = 'AzureADMyOrg'
}

$app = New-MgApplication @params

if ($AppType -ieq 'Set-OutlookSignatures') {
    Write-Host "  App Client ID for Set-OutlookSignatures graph config file: $($app.AppId)" -ForegroundColor Green
} else {
    Write-Host "  App Client ID for SimulateAndDeploy configuration: $($app.AppId)" -ForegroundColor Green
}


Write-Host
Write-Host 'Add required permissions to app registration'
if ($AppType -ieq 'Set-OutlookSignatures') {
    $params = @{
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
} else {
    $params = @{
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
                        'id'   = 'df85f4d6-205c-4ac5-a5ea-6bf408dba283'
                        'type' = 'Scope'
                    },

                    # Application permission: GroupMember.Read.All
                    #   Allows the app to list groups, read basic group properties and read membership of all groups the signed-in user has access to.
                    #   Required to find groups by name and to get their security identifier (SID) and the number of transitive members.
                    @{
                        'id'   = '98830695-27a2-44f7-8c18-0c3ebc9698f6'
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
}

Update-MgApplication -ApplicationId $app.Id -BodyParameter $params


Write-Host
Write-Host 'Add redirect URIs to app registration'
$params =	@{
    RedirectUris = @(
        'http://localhost',
        "ms-appx-web://microsoft.aad.brokerplugin/$($app.AppId)"
    )
}

Update-MgApplication -ApplicationId $app.Id -IsFallbackPublicClient -PublicClient $params


Write-Host
Write-Host 'Enable public client flow'
Update-MgApplication -ApplicationId $app.Id -IsFallbackPublicClient


if ($AppType -ieq 'SimulateAndDeploy') {
    Write-Host
    Write-Host 'Add client secret to app registration'

    $params = @{
        displayName = "Initial client secret, valid $(Get-Date -Format 'yyyy-MM-dd')--$(Get-Date (Get-Date).AddMonths(24) -Format 'yyyy-MM-dd')"
        endDateTime = (Get-Date).AddMonths(24)
    }

    $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $params

    Write-Host "  Client secret for SimulateAndDeploy configuration: $($secret.SecretText)" -ForegroundColor Green
    Write-Host "  Don't forget to renew the client secret before $(Get-Date (Get-Date).AddMonths(24) -Format 'yyyy-MM-dd')" -ForegroundColor Green
}


Write-Host
Write-Host 'Consider restricting file access'
Write-Host '  Consider switching from Files.Read.All to Files.SelectedOperations.Selected for added security.'
Write-Host '    This requires granting specific permissions in SharePoint Online.'


Write-Host
Write-Host 'Grant admin consent'
Write-Host ('  This creates an enterprise application from the app registration and makes the app accessible to ' + $(
        if ($AppType -ieq 'Set-OutlookSignatures') {
            Write-Host 'end users running Set-OutlookSignatures'
        } else {
            Write-Host 'the account running Set-OutlookSignatures in SimulateAndDeploy mode'
        }
    )
)
Write-Host '  To grant admin consent, navigate to'
Write-Host "    https://login.microsoftonline.com/$($app.PublisherDomain)/adminconsent?client_id=$($app.AppId)" -ForegroundColor Green
Write-Host '    with a user being 'Application Adminstrator' or 'Global Administrator' and accept the required permissions on behalf of your tenant.'
Write-Host "  You can safely ignore the error message that the URL 'http://localhost/?admin_consent=True&tenant=[…]'"
Write-Host '    could not be found or accessed. The reason for this message is that the Entra ID app is configured to only be able to authenticate against http://localhost.'


Write-Host
Write-Host 'Done'
