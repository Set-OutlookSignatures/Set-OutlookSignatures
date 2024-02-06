# This file allows defining the default configuration for connecting to Microsoft Graph for Set-OutlookSignatures
#
# This script is executed as a whole once per Set-OutlookSignatures run.
#
# Attention: The configuration file is executed as part of Set-OutlookSignatures.ps1 and is not checked for any harmful content. Please only allow qualified technicians write access to this file, only use it to to define replacement variables and test it thoroughly.
#
# A variable defined in this file overrides the definition of the same variable defined earlier in the software.
#
#
# What is the recommended approach for custom configuration files?
# You should not change the default configuration file '.\config\default graph config.ps1', as it might be changed in a future release of Set-OutlookSignatures. In this case, you would have to sort out the changes yourself.
#
# The following steps are recommended:
# 1. Create a new custom configuration file in a separate folder.
# 2. The first step in the new custom configuration file should be to load the default configuration file:
#    # Loading default replacement variables shipped with Set-OutlookSignatures
#    . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $(Join-Path -Path $(Get-Location).ProviderPath -ChildPath '\config\default graph config.ps1') -Raw)))
# 3. After importing the default configuration file, existing configurations and mappings can be altered with custom definitions and new ones can be added.
# 4. Instead of altering existing replacement variables, it is recommended to create new replacement variables with modified content.
# 5. Start Set-OutlookSignatures with the parameter 'GraphConfigFile' pointing to the new custom configuration file.


# Client ID
# The default client ID is defined in the developers Entra ID/Azure AD tenant as multi-tenant, so it can be used everywhere
#   For security and maintenance reasons, it is recommended to create you own app in your own tenant
# It can be replaced with the ID of an app created in your own tenant
#   Option A: Create the app automatically by using the script '.\sample code\Create-EntraApp.ps1'
#     The sample code creates the app with all required settings automatically, only providing admin consent is a manual task
#   Option B: Create the Entra app manually
#     Create an app in Entra admin center (https://entra.microsoft.com)
#       Sign in as at least Cloud Application Administrator
#       Identity > Applications > App registrations > New registration
#       Enter at least a display name for your application
#       Set "Supported account type" to "Accounts in this organizational directory only"
#       Set Redirect URI to "Mobile and desktop applications" and 'http://localhost' (http, not https)
#       The "Application (client) ID" is the value you need to set for $GraphClientID in this file
#     Client secret
#       There is no need to define a client secret, as we only work with delegated permissions, and not with application permissions
#     Add the following delegated permissions (not application permissions)
#       Identity > Applications > App registrations > your application > API permissions > Add a permission
#       Microsoft Graph
#         email
#           Allows the app to read your users' primary email address.
#           Required to log on the current user.
#         EWS.AccessAsUser.All
#           Allows the app to have the same access to mailboxes as the signed-in user via Exchange Web Services.
#           Required to connect to Outlook Web and to set Outlook Web signature (classic and roaming).
#         GroupMember.Read.All
#           Allows the app to list groups, read basic group properties and read membership of all groups the signed-in user has access to.
#           Required to find groups by name and to get their security identifier (SID) and the number of transitive members.
#         MailboxSettings.ReadWrite
#           Allows the app to create, read, update, and delete user's mailbox settings. Does not include permission to send mail.
#           Required to detect the state of the out of office assistant and to set out-of-office replies.
#         offline_access
#           Allows the app to see and update the data you gave it access to, even when users are not currently using the app. This does not give the app any additional permissions.
#           Required to get a refresh token from Graph.
#         openid
#           Allows users to sign in to the app with their work or school accounts and allows the app to see basic user profile information.
#           Required to log on the current user.
#         profile
#           Allows the app to see your users' basic profile (e.g., name, picture, user name, email address).
#           Required to log on the current user, to access the '/me' Graph API, to get basic properties of the current user.
#         User.Read.All
#           Allows the app to read the full set of profile properties, reports, and managers of other users in your organization, on behalf of the signed-in user.
#           Required for $CurrentUser[...]$ and $CurrentMailbox[...]$ replacement variables, and for simulation mode.
#       Provide admin consent
#         Click the "Grant admin consent for {your tenant}" button
#     Enable 'Allow public client flows'
#       Identity > Applications > App registrations > your application > Advanced settings
#       Enable "Allow public client flows"
#       This enables SSO (single sign-on) for domain-joined Windows (Integrated Windows Authentication Flow)
$GraphClientID = 'beea8249-8c98-4c76-92f6-ce3c468a61e6'


# Endpoint version
$GraphEndpointVersion = 'v1.0'


# Message box text to show before browser opens for authentication to Microsoft 365
# Leave blank to not show message box at all
# Defining a text is recommended to inform users about the upcoming opening of a new browser tab asking for authentication
#   Set-OutlookSignatures usually runs in the background and the M365 logon screen can not show a hint to Set-OutlookSignatures,
#   so the new tab is a negative surprise for users
# The message box is shown on the very top of the active desktop and can not be sent to the background,
#   but does not steal the focus
$GraphHtmlMessageboxText = "You started Set-OutlookSignatures, or an administrator configured it to run for you to update your Outlook signatures and out-of-office replies.`r`n`r`nA required token for access to Microsoft 365 is not yet available.`r`n`The program may run for the first time on this client, a previous token may have expired or been deleted.`r`n`r`nTo create this required token, please login to Microsoft 365 with your account ""$($script:CurrentUser)"" in the new browser tab that will open after you close this message."


# HTML message to show after successful browser authentication to Microsoft Graph
# Try to keep the resulting HTML code small, as long code may lead to display errors ("connection reset")
$GraphHtmlMessageSuccess = "<html><head><title>$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })</title></head><body style=""font-family:sans-serif;""><h1><a href=""$(if ($BenefactorCircleLicenseFile) { 'https://explicitconsulting.at/open-source/set-outlooksignatures' } else { 'https://github.com/Set-OutlookSignatures/Set-OutlookSignatures' })"" target=""_blank"">$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })</a></h1><p>Graph authentication successful at $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK').</p><p>&nbsp;</p><p><strong>Thank you for using <a href=""$(if ($BenefactorCircleLicenseFile) { 'https://explicitconsulting.at/open-source/set-outlooksignatures' } else { 'https://github.com/Set-OutlookSignatures/Set-OutlookSignatures' })"" target=""_blank"">$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })</a>!</strong></p><p>&nbsp;</p><p>You can close this tab at any time.</p></body></html>"
# Example with automatically closing tab: $GraphHtmlMessageSuccess = "<html><head><script>window.close();</script><title>$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })</title></head><body style=""font-family:sans-serif;""><h1><a href=""$(if ($BenefactorCircleLicenseFile) { 'https://explicitconsulting.at/open-source/set-outlooksignatures' } else { 'https://github.com/Set-OutlookSignatures/Set-OutlookSignatures' })"" target=""_blank"">$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })</a></h1><p>Graph authentication successful at $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK').</p><p>&nbsp;</p><p><strong>Thank you for using <a href=""$(if ($BenefactorCircleLicenseFile) { 'https://explicitconsulting.at/open-source/set-outlooksignatures' } else { 'https://github.com/Set-OutlookSignatures/Set-OutlookSignatures' })"" target=""_blank"">$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })</a>!</strong></p><p>&nbsp;</p><p>You can close this tab at any time.</p></body></html>"


# When the user successfully authenticates in the browser, the browser will be redirected to to the given Uri
# Takes precedence over $GraphHtmlMessageSuccess
[uri] $GraphBrowserRedirectSuccess = ''


# HTML message to show after unsuccessful browser authentication to Microsoft Graph
# Try to keep the resulting HTML code small, as long code may lead to display errors ("connection reset")
$GraphHtmlMessageError = "<html><head><title>$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })</title></head><body style=""font-family:sans-serif;""><h1><a href=""$(if ($BenefactorCircleLicenseFile) { 'https://explicitconsulting.at/open-source/set-outlooksignatures' } else { 'https://github.com/Set-OutlookSignatures/Set-OutlookSignatures' })"" target=""_blank"">$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })</a></h1><p>Graph authentication failed at $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK').</p><p>You may want to inform your helpdesk or administrator about this problem.</p><p>&nbsp;</p><p>Error: {0}</p><p>Details: {1}</p><p>&nbsp;</p><p><strong>Thank you for using <a href=""$(if ($BenefactorCircleLicenseFile) { 'https://explicitconsulting.at/open-source/set-outlooksignatures' } else { 'https://github.com/Set-OutlookSignatures/Set-OutlookSignatures' })"" target=""_blank"">$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })</a>!</strong></p></body></html>"


# When the user fails to successfully authenticate in the browser, the browser will be redirected to to the given Uri
# Takes precedence over $GraphHtmlMessageError
[uri] $GraphBrowserRedirectError = ''


# User properties to select
# Custom Graph attributes: 'extension_<AppID owning the extension attribute>_<attribute name>'
$GraphUserProperties = @(
    'aboutMe',
    'businessPhones',
    'city',
    'companyName',
    'country',
    'department',
    'displayName',
    'faxNumber',
    'givenName',
    'jobTitle',
    'mail',
    'mailNickname',
    'mobilePhone',
    'officeLocation',
    'onPremisesDistinguishedName',
    'onPremisesDomainName',
    'onPremisesExtensionAttributes',
    'onPremisesImmutableId',
    'onPremisesSamAccountName',
    'onPremisesSecurityIdentifier',
    'onPremisesUserPrincipalName',
    'postalCode',
    'proxyAddresses',
    'state',
    'streetAddress',
    'surname',
    'usageLocation',
    'userPrincipalName'
)


# Mapping Graph user properties to on-prem Active Directory user properties
# This way, we do not need to differentiate between on-prem, hybrid and cloud in '.\config\default replacement variables.ps1'
# Active Directory attribute names on the left, Graph attribute names on the right
# Custom Graph attributes: 'extension_<AppID owning the extension attribute>_<attribute name>'
$GraphUserAttributeMapping = @{
    co                         = 'country'
    company                    = 'companyName'
    department                 = 'department'
    displayName                = 'displayName'
    extensionAttribute1        = 'onPremisesExtensionAttributes.extensionAttribute1'
    extensionAttribute2        = 'onPremisesExtensionAttributes.extensionAttribute2'
    extensionAttribute3        = 'onPremisesExtensionAttributes.extensionAttribute3'
    extensionAttribute4        = 'onPremisesExtensionAttributes.extensionAttribute4'
    extensionAttribute5        = 'onPremisesExtensionAttributes.extensionAttribute5'
    extensionAttribute6        = 'onPremisesExtensionAttributes.extensionAttribute6'
    extensionAttribute7        = 'onPremisesExtensionAttributes.extensionAttribute7'
    extensionAttribute8        = 'onPremisesExtensionAttributes.extensionAttribute8'
    extensionAttribute9        = 'onPremisesExtensionAttributes.extensionAttribute9'
    extensionAttribute10       = 'onPremisesExtensionAttributes.extensionAttribute10'
    extensionAttribute11       = 'onPremisesExtensionAttributes.extensionAttribute11'
    extensionAttribute12       = 'onPremisesExtensionAttributes.extensionAttribute12'
    extensionAttribute13       = 'onPremisesExtensionAttributes.extensionAttribute13'
    extensionAttribute14       = 'onPremisesExtensionAttributes.extensionAttribute14'
    extensionAttribute15       = 'onPremisesExtensionAttributes.extensionAttribute15'
    facsimileTelephoneNumber   = 'faxNumber'
    givenName                  = 'givenName'
    l                          = 'city'
    mail                       = 'mail'
    mailNickname               = 'mailNickname'
    mobile                     = 'mobilePhone'
    physicalDeliveryOfficeName = 'officeLocation'
    postalCode                 = 'postalCode'
    proxyAddresses             = 'proxyAddresses'
    sn                         = 'surname'
    st                         = 'state'
    streetAddress              = 'streetAddress'
    telephoneNumber            = 'businessPhones'
    title                      = 'jobTitle'
}
