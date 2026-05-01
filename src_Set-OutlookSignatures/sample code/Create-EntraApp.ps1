<#
This sample code shows how to automate the creation of the Entra ID app required for Set-OutlookSignatures.

Both types of apps are supported: The one for end users, and the one for SimulateAndDeploy.

You can adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers professional support for this and other open source code.
#>


[CmdletBinding()]

param (
    # Which type of app should be created?
    #   'Set-OutlookSignatures' for the default Set-OutlookSignatures app being accessed by end users runnding Set-OutlookSignatures
    #     Uses only delegated permissions, as described in '.\config\default graph config.ps1'
    #   'SimulateAndDeploy' for use in the "simulate and deploy" scenario
    #     Uses delegated permissions and application permissions, as described in '.\sample code\SimulateAndDeploy.ps1'
    #   For security reasons, the app type has no default value and needs to be set manually
    [ValidateSet('Set-OutlookSignatures', 'SimulateAndDeploy', 'OutlookAddIn', IgnoreCase = $true)]
    $AppType = $null,

    # Name of the Entra ID application to create
    [ValidateNotNullOrEmpty()]
    $AppName = $null,

    # Outlook add-in url to be used in the Entra ID application
    [ValidateNotNullOrEmpty()]
    [uri]$OutlookAddInUrl = $null,

    # Cloud environment to use
    # Built-in values: 'Public', 'Global', 'AzurePublic', 'AzureGlobal', 'AzureCloud', 'AzureUSGovernmentGCC', 'USGovernmentGCC', 'AzureUSGovernment', 'AzureUSGovernmentGCCHigh', 'AzureUSGovernmentL4', 'USGovernmentGCCHigh', 'USGovernmentL4', 'AzureUSGovernmentDOD', 'AzureUSGovernmentL5', 'USGovernmentDOD', 'USGovernmentL5', 'China', 'AzureChina', 'ChinaCloud', 'AzureChinaCloud', 'Bleu', 'AzureBleu', 'BleuCloud', 'AzureBleuCloud', 'Delos', 'AzureDelos', 'DelosCloud', 'AzureDelosCloud', 'GovSG', 'AzureGovSG', 'GovSGCloud', 'AzureGovSGCloud'
    # Other values require defining $MgGraphAzureADEndpoint and $MgGraphGraphEndpoint
    [ValidateNotNullOrEmpty()]
    [string]$CloudEnvironment = 'Public',

    # String for AzureADEndpoint parameter to be used with Connect-MgGraph
    # Only required for non built-in values for $CloudEnvironment
    [string]$MgGraphAzureADEndpoint = $null, # Example: 'https://login.sovcloud-identity.example/'

    # String for GraphEndpoint parameter to be used with Connect-MgGraph
    # Only required for non built-in values for $CloudEnvironment
    [string]$MgGraphGraphEndpoint = $null, # Example: 'https://graph.svc.sovcloud.example/'

    # Application ID of the app to use when connecting with Connect-MgGraph
    [string]$MgGraphAppClientId = $null
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
    Write-Host '  All apps require the AppType and AppName parameters, app type OutlookAddIn additionally the OutlookAddInUrl parameter.' -ForegroundColor Red

    exit 1
} else {
    $AppName = $AppName.trim()
}

switch ($CloudEnvironment) {
    { $_ -iin @('Public', 'Global', 'AzurePublic', 'AzureGlobal', 'AzureCloud', 'AzureUSGovernmentGCC', 'USGovernmentGCC') } {
        $MgGraphEnvironment = 'Global'
        break
    }

    { $_ -iin @('AzureUSGovernment', 'AzureUSGovernmentGCCHigh', 'AzureUSGovernmentL4', 'USGovernmentGCCHigh', 'USGovernmentL4') } {
        $MgGraphEnvironment = 'USGov'
        break
    }

    { $_ -iin @('AzureUSGovernmentDOD', 'AzureUSGovernmentL5', 'USGovernmentDOD', 'USGovernmentL5') } {
        $MgGraphEnvironment = 'USGovDoD'
        break
    }

    { $_ -iin @('China', 'AzureChina', 'ChinaCloud', 'AzureChinaCloud') } {
        $MgGraphEnvironment = 'China'
        break
    }

    { $_ -iin @('Bleu', 'AzureBleu', 'BleuCloud', 'AzureBleuCloud') } {
        $MgGraphEnvironment = 'BleuCloud'
        $MgGraphAzureADEndpoint = 'https://login.sovcloud-identity.fr/'
        $MgGraphGraphEndpoint = 'https://graph.svc.sovcloud.fr/'
        break
    }

    { $_ -iin @('Delos', 'AzureDelos', 'DelosCloud', 'AzureDelosCloud') } {
        $MgGraphEnvironment = 'DelosCloud'
        $MgGraphAzureADEndpoint = 'https://login.sovcloud-identity.de/'
        $MgGraphGraphEndpoint = 'https://graph.svc.sovcloud.de/'
        break
    }

    { $_ -iin @('GovSG', 'AzureGovSG', 'GovSGCloud', 'AzureGovSGCloud') } {
        $MgGraphEnvironment = 'GovSGCloud'
        $MgGraphAzureADEndpoint = 'https://login.sovcloud-identity.sg/'
        $MgGraphGraphEndpoint = 'https://graph.svc.sovcloud.sg/'
        break
    }

    default {
        $MgGraphEnvironment = $CloudEnvironment
        break
    }
}


Write-Host
Write-Host 'Entra ID app to create'
Write-Host "  App type: $($AppType)"
Write-Host "  App name: $($AppName)"
if ($AppType -ieq 'OutlookAddIn') {
    Write-Host "  Outlook Add-In URI: $($OutlookAddInUrl)"
}


Write-Host
Write-Host 'Install required PowerShell modules'
[enum]::GetNames([System.Net.SecurityProtocolType]) | ForEach-Object {
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol, $_
    } catch {
    }
}

try {
    if (-not (Get-PSRepository | Where-Object { $_.Name -ieq 'PSGallery' })) {
        Register-PSRepository -Name 'PSGallery' -SourceLocation 'https://www.powershellgallery.com/api/v2' -WarningAction SilentlyContinue -ErrorAction Stop
    }

    @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications', 'Microsoft.Graph.Identity.SignIns') | ForEach-Object {
        Write-Host "  $($_)"

        if (Get-Module -ListAvailable -Name $_) {
            Find-Module -Name $_ -Repository PSGallery | Update-Module -Force -WarningAction SilentlyContinue -ErrorAction Stop
        } else {
            Find-Module -Name $_ -Repository PSGallery | Install-Module -Force -AllowClobber -WarningAction SilentlyContinue -ErrorAction Stop
        }

        Import-Module -Name $_ -Force -WarningAction SilentlyContinue -ErrorAction Stop
    }
} catch {
    Write-Host "Error installing PowerShell modules: $($_)" -ForegroundColor Red
    Write-Host
    Write-Host 'This is a severe error. It is not related to this script, but to the basic PowerShell setup on this system.' -ForegroundColor Red
    Write-Host 'Please fix these issues with PowerShell package management, package providers and modules first.' -ForegroundColor Red

    exit 1
}


if ((Get-MgEnvironment).Name -inotcontains $MgGraphEnvironment) {
    Write-Host
    Write-Host "Adding custom cloud environment '$($MgGraphEnvironment)'"

    if ([String]::IsNullOrEmpty($MgGraphAzureADEndpoint) -or [String]::IsNullOrEmpty($MgGraphGraphEndpoint)) {
        Write-Host "  '$($MgGraphEnvironment)' is a custom environment, so `$MgGraphAzureADEndpoint and `$MgGraphGraphEndpoint must be set." -ForegroundColor Red
        exit 1
    }

    Add-MgEnvironment -Name $MgGraphEnvironment -AzureAdEndpoint $MgGraphAzureADEndpoint -GraphEndpoint $MgGraphGraphEndpoint
}


Write-Host
Write-Host "Connect to your Entra ID with a user being 'Application Administrator' or 'Global Administrator'"
Write-Host "  Connecting to Graph environment '$($MgGraphEnvironment)'"
Write-Host "    To connect to another environment, cancel authentication and add the '-CloudEnvironment' parameter."
Write-Host '  An authentication window will open, likely in a browser'

# Disconnect first, so that no existing connection is re-used. This forces to choose an account for the following connect.
$null = Disconnect-MgGraph -ErrorAction SilentlyContinue

try {
    $scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All')

    if ($MgGraphEnvironment -iin @('BleuCloud', 'DelosCloud', 'GovSGCloud')) {
        Connect-MgGraph -Environment $MgGraphEnvironment -ClientId $MgGraphAppClientId -ContextScope Process -Scopes $scopes -NoWelcome -ErrorAction Stop
    } else {
        Connect-MgGraph -Environment $MgGraphEnvironment -ContextScope Process -Scopes $scopes -NoWelcome -ErrorAction Stop
    }

    if (-not (Get-MgContext)) {
        throw 'No connection established.'
    } else {
        $scopes | ForEach-Object {
            if (-not (Get-MgContext).Scopes -icontains ($_)) {
                throw "Required scope '$_' not granted."
            }
        }
    }
} catch {
    Write-Host "Error connecting to Microsoft Graph: $($_)" -ForegroundColor Red
    Write-Host
    Write-Host 'Please ensure that you can connect to Microsoft Graph and that your user has sufficient permissions.' -ForegroundColor Red

    exit 1
}

Write-Host
Write-Host 'Create a new app registration'
Write-Host "  App name: $($AppName)"

$ExistingApp = @(Get-MgApplication -Filter "DisplayName eq '$($AppName)'" -ErrorAction Stop)

if ($ExistingApp.Count -gt 0) {
    $ExistingApp | ForEach-Object {
        Write-Host "  App with name '$($AppName)' already exists. ID: $($_.Id)" -ForegroundColor Red
    }

    Write-Host '  Exiting.' -ForegroundColor Red
    exit 1
}

$params = @{
    DisplayName    = $AppName
    Description    = "$($AppType) app for Set-OutlookSignatures: Data Sovereign Email Signatures and Out-of-Office Replies"
    Notes          = "$($AppType) app for Set-OutlookSignatures: Data Sovereign Email Signatures and Out-of-Office Replies"
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
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#email
                    #   Authenticate the signed-in user.
                    @{
                        'id'   = '64a6cdd6-aab1-4aaf-94b8-3cc8405e90d0'
                        'type' = 'Scope'
                    },

                    # Delegated permission: MailboxConfigItem.ReadWrite
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxconfigitemreadwrite
                    #   Read data from Outlook Web, set Outlook web signatures.
                    @{
                        'id'   = '7d461784-7715-4b09-9f90-91a6d8722652'
                        'type' = 'Scope'
                    },

                    # Delegated permission: Files.Read.All
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#filesreadall
                    #   Read template and configuration files hosted on SharePoint Online. Alternative: Files.SelectedOperations.Selected.
                    @{
                        'id'   = 'df85f4d6-205c-4ac5-a5ea-6bf408dba283'
                        'type' = 'Scope'
                    },

                    # Delegated permission: GroupMember.Read.All
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#groupmemberreadall
                    #   Find groups by name, get their security identifier (SID) and transitive members.
                    @{
                        'id'   = 'bc024368-1153-4739-b217-4326f2e966d0'
                        'type' = 'Scope'
                    },

                    # Delegated permission: Mail.ReadWrite
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#mailreadwrite
                    #   Create signature collection in drafts, provide signatures for Outlook add-in.
                    @{
                        'id'   = '024d486e-b451-40bb-833d-3e66d98c5c73'
                        'type' = 'Scope'
                    },

                    # Delegated permission: MailboxSettings.ReadWrite
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxsettingsreadwrite
                    #   Detect mailbox environment, get and set out-of-office data.
                    @{
                        'id'   = '818c620a-27a9-40bd-a6a5-d96f7d610b4b'
                        'type' = 'Scope'
                    },

                    # Delegated permission: offline_access
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#offline_access
                    #   Required to get a refresh token from Graph.
                    @{
                        'id'   = '7427e0e9-2fba-42fe-b0c0-848c9e6a8182'
                        'type' = 'Scope'
                    },

                    # Delegated permission: openid
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#openid
                    #   Authenticate the signed-in user.
                    @{
                        'id'   = '37f7f235-527c-4136-accd-4a02d197296e'
                        'type' = 'Scope'
                    },

                    # Delegated permission: profile
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#profile
                    #   Authenticate the signed-in user, get basic properties.
                    @{
                        'id'   = '14dad69e-099b-42c9-810b-d002981feec1'
                        'type' = 'Scope'
                    },

                    # Delegated permission: User.Read.All
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#userreadall
                    #   Data for replacement variables, SMTP to UPN, group membership.
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

                    # Delegated permission: email
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#email
                    #   Authenticate the signed-in user.
                    @{
                        'id'   = '64a6cdd6-aab1-4aaf-94b8-3cc8405e90d0'
                        'type' = 'Scope'
                    },

                    # Delegated permission: MailboxConfigItem.ReadWrite
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxconfigitemreadwrite
                    #   Read data from Outlook Web, set Outlook web signatures.
                    @{
                        'id'   = '7d461784-7715-4b09-9f90-91a6d8722652'
                        'type' = 'Scope'
                    },

                    # Delegated permission: Files.Read.All
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#filesreadall
                    #   Read template and configuration files hosted on SharePoint Online. Alternative: Files.SelectedOperations.Selected.
                    @{
                        'id'   = 'df85f4d6-205c-4ac5-a5ea-6bf408dba283'
                        'type' = 'Scope'
                    },

                    # Delegated permission: GroupMember.Read.All
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#groupmemberreadall
                    #   Find groups by name, get their security identifier (SID) and transitive members.
                    @{
                        'id'   = 'bc024368-1153-4739-b217-4326f2e966d0'
                        'type' = 'Scope'
                    },

                    # Delegated permission: Mail.ReadWrite
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#mailreadwrite
                    #   Create signature collection in drafts, provide signatures for Outlook add-in.
                    @{
                        'id'   = '024d486e-b451-40bb-833d-3e66d98c5c73'
                        'type' = 'Scope'
                    },

                    # Delegated permission: MailboxSettings.ReadWrite
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxsettingsreadwrite
                    #   Detect mailbox environment, get and set out-of-office data.
                    @{
                        'id'   = '818c620a-27a9-40bd-a6a5-d96f7d610b4b'
                        'type' = 'Scope'
                    },

                    # Delegated permission: offline_access
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#offline_access
                    #   Required to get a refresh token from Graph.
                    @{
                        'id'   = '7427e0e9-2fba-42fe-b0c0-848c9e6a8182'
                        'type' = 'Scope'
                    },

                    # Delegated permission: openid
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#openid
                    #   Authenticate the signed-in user.
                    @{
                        'id'   = '37f7f235-527c-4136-accd-4a02d197296e'
                        'type' = 'Scope'
                    },

                    # Delegated permission: profile
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#profile
                    #   Authenticate the signed-in user, get basic properties.
                    @{
                        'id'   = '14dad69e-099b-42c9-810b-d002981feec1'
                        'type' = 'Scope'
                    },

                    # Delegated permission: User.Read.All
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#userreadall
                    #   Data for replacement variables, SMTP to UPN, group membership.
                    @{
                        'id'   = 'a154be20-db9c-4678-8ab7-66f6cc099a59'
                        'type' = 'Scope'
                    },

                    # Application permission: Files.Read.All
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#filesreadall
                    #   Read template and configuration files hosted on SharePoint Online. Alternative: Files.SelectedOperations.Selected.
                    @{
                        'id'   = '01d4889c-1287-42c6-ac1f-5d1e02578ef6'
                        'type' = 'Role'
                    },

                    # Application permission: GroupMember.Read.All
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#groupmemberreadall
                    #   Find groups by name, get their security identifier (SID) and transitive members.
                    @{
                        'id'   = '98830695-27a2-44f7-8c18-0c3ebc9698f6'
                        'type' = 'Role'
                    },

                    # Application permission: Mail.ReadWrite
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#mailreadwrite
                    #   Create signature collection in drafts, provide signatures for Outlook add-in.
                    @{
                        'id'   = 'e2a3a72e-5f79-4c64-b1b1-878b674786c9'
                        'type' = 'Role'
                    },

                    # Application permission: MailboxSettings.ReadWrite
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxsettingsreadwrite
                    #   Detect mailbox environment, get and set out-of-office data.
                    @{
                        'id'   = '6931bccd-447a-43d1-b442-00a195474933'
                        'type' = 'Role'
                    },

                    # Application permission: User.Read.All
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#userreadall
                    #   Data for replacement variables, SMTP to UPN, group membership.
                    @{
                        'id'   = 'df021288-bdef-4463-88db-98f22de89214'
                        'type' = 'Role'
                    },

                    # Application permission: MailboxConfigItem.ReadWrite
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxconfigitemreadwrite
                    #   Read data from Outlook Web, set Outlook web signatures.
                    @{
                        'id'   = 'aa6d92d4-b25a-4640-aefe-3e3231e5e736'
                        'type' = 'Role'
                    }
                )
            }
        )
    }
} elseif ($AppType -ieq 'OutlookAddIn') {
    $permissionParams = @{
        RequiredResourceAccess = @(
            @{
                # Microsoft Graph
                'resourceAppId'  = '00000003-0000-0000-c000-000000000000'
                'resourceAccess' = @(
                    # Microsoft Graph permissions reference: https://learn.microsoft.com/en-us/graph/permissions-reference

                    # Delegated permission: GroupMember.Read.All
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#groupmemberreadall
                    #   Find groups by name, get their security identifier (SID) and transitive members.
                    @{
                        'id'   = 'bc024368-1153-4739-b217-4326f2e966d0'
                        'type' = 'Scope'
                    },

                    # Delegated permission: Mail.Read
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#mailread
                    #   Required because of Microsoft restrictions accessing roaming signatures.
                    @{
                        'id'   = '570282fd-fa5c-430d-a7fd-fc8dc98a9dca'
                        'type' = 'Scope'
                    },

                    # Delegated permission: User.Read.All
                    #   https://learn.microsoft.com/en-us/graph/permissions-reference#userreadall
                    #   Data for replacement variables, SMTP to UPN, group membership.
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
Write-Host '  This may take a moment'
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
Write-Host 'Relevant information for your configuration below' -ForegroundColor Green
if ($AppType -ieq 'Set-OutlookSignatures') {
    Write-Host "  GraphClientId for Set-OutlookSignatures: '$($app.AppId)'"
} elseif ($AppType -ieq 'SimulateAndDeploy') {
    Write-Host "  GraphClientId for SimulateAndDeploy: '$($app.AppId)'"
    Write-Host "  GraphClientSecret for SimulateAndDeploy: '$($secret.SecretText)'"
    Write-Host "    Do not forget to renew the client secret before $(Get-Date (Get-Date).AddMonths(24) -Format 'yyyy-MM-dd')!"
} elseif ($AppType -ieq 'OutlookAddIn') {
    Write-Host "  GRAPH_CLIENT_ID for Outlook Add-In: '$($app.AppId)'"
}
Write-Host 'Relevant information for your configuration above' -ForegroundColor Green

Write-Host
Write-Host 'Done'