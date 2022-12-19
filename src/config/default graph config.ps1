# This file allows defining the default configuration for connecting to Microsoft Graph for Set-OutlookSignatures
#
# This script is executed as a whole once per Set-OutlookSignatures run.
#
# Attention: The configuration file is executed as part of Set-OutlookSignatures.ps1 and is not checked for any harmful content. Please only allow qualified technicians write access to this file, only use it to to define replacement variables and test it thoroughly.
#
# Active Directory property names are case sensitive.
# It is required to use full lowercase Active Directory property names.
#
# A variable defined in this file overrides the definition of the same variable defined earlier in the script.
#
#
# What is the recommended approach for custom configuration files?
# You should not change the default configuration file `'.\config\default graph config.ps1'`, as it might be changed in a future release of Set-OutlookSignatures. In this case, you would have to sort out the changes yourself.
#
# The following steps are recommended:
# 1. Create a new custom configuration file in a separate folder.
# 2. The first step in the new custom configuration file should be to load the default configuration file:
#    # Loading default replacement variables shipped with Set-OutlookSignatures
#    . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $(Join-Path -Path $(Get-Location).path -ChildPath '\config\default graph config.ps1') -Raw)))
# 3. After importing the default configuration file, existing configurations and mappings can be altered with custom definitions and new ones can be added.
# 4. Instead of altering existing replacement variables, it is recommended to create new replacement variables with modified content.
# 5. Start Set-OutlookSignatures with the parameter 'GraphConfigFile' pointing to the new custom configuration file.


# Client ID
# The default client ID is defined in the developer's Azure tenant as multi-tenant, so it can be used everywhere
# Can be replaced with a Client ID from the own tenant
#   Scopes (please provide admin consent): 'https://graph.microsoft.com/openid', 'https://graph.microsoft.com/email', 'https://graph.microsoft.com/profile', 'https://graph.microsoft.com/user.read.all', 'https://graph.microsoft.com/group.read.all', 'https://graph.microsoft.com/mailboxsettings.readwrite', 'https://graph.microsoft.com/EWS.AccessAsUser.All'
#   Redirect URI: 'http://localhost', configure for "mobile and desktop applications"
#   Enable 'Allow public client flows' to make Windows Integrated Authentication (SSO) work for Azure AD joined devices
$GraphClientID = 'beea8249-8c98-4c76-92f6-ce3c468a61e6'


# Endpoint version
$GraphEndpointVersion = 'v1.0'


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
    givenname                  = 'givenName'
    sn                         = 'surname'
    department                 = 'department'
    title                      = 'jobTitle'
    streetaddress              = 'streetAddress'
    postalcode                 = 'postalCode'
    l                          = 'city'
    co                         = 'country'
    telephonenumber            = 'businessPhones'
    facsimiletelephonenumber   = 'faxNumber'
    mobile                     = 'mobilePhone'
    mail                       = 'mail'
    extensionattribute1        = 'onPremisesExtensionAttributes.extensionAttribute1'
    extensionattribute2        = 'onPremisesExtensionAttributes.extensionAttribute2'
    extensionattribute3        = 'onPremisesExtensionAttributes.extensionAttribute3'
    extensionattribute4        = 'onPremisesExtensionAttributes.extensionAttribute4'
    extensionattribute5        = 'onPremisesExtensionAttributes.extensionAttribute5'
    extensionattribute6        = 'onPremisesExtensionAttributes.extensionAttribute6'
    extensionattribute7        = 'onPremisesExtensionAttributes.extensionAttribute7'
    extensionattribute8        = 'onPremisesExtensionAttributes.extensionAttribute8'
    extensionattribute9        = 'onPremisesExtensionAttributes.extensionAttribute9'
    extensionattribute10       = 'onPremisesExtensionAttributes.extensionAttribute10'
    extensionattribute11       = 'onPremisesExtensionAttributes.extensionAttribute11'
    extensionattribute12       = 'onPremisesExtensionAttributes.extensionAttribute12'
    extensionattribute13       = 'onPremisesExtensionAttributes.extensionAttribute13'
    extensionattribute14       = 'onPremisesExtensionAttributes.extensionAttribute14'
    extensionattribute15       = 'onPremisesExtensionAttributes.extensionAttribute15'
    objectsid                  = 'onPremisesSecurityIdentifier'
    distinguishedname          = 'onPremisesDistinguishedName'
    company                    = 'companyName'
    displayname                = 'displayName'
    proxyAddresses             = 'proxyAddresses'
    userprincipalname          = 'userPrincipalName'
    physicaldeliveryofficename = 'officeLocation'
    mailboxsettings            = 'mailboxsettings'
    mailnickname               = 'mailNickname'
    st                         = 'state'
}
