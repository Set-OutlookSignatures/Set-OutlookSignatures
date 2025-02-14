# This file allows defining custom replacement variables for Set-OutlookSignatures
#
# This script is executed as a whole once for each mailbox.
# It allows for complex replacement variable handling (complex string transformations, retrieving information from web services and databases, etc.).
#
# Attention: The configuration file is executed as part of Set-OutlookSignatures.ps1 and is not checked for any harmful content. Please only allow qualified technicians write access to this file, only use it to to define replacement variables and test it thoroughly.
#
# Replacement variable names are not case sensitive.
#
# A variable defined in this file overrides the definition of the same variable defined earlier in the software.
#
#
# See README file for more examples, such as:
#   Allowed tags
#   How to work with ini files
#   Variable replacement
#   Photos from Active Directory
#   Delete images when attribute is empty, variable content based on group membership
#   How to avoid blank lines when replacement variables return an empty string?
#
#
# What is the recommended approach for custom configuration files?
# You should not change the default configuration file '.\config\default replacement variable.ps1', as it might be changed in a future release of Set-OutlookSignatures. In this case, you would have to sort out the changes yourself.
#
# The following steps are recommended:
# 1. Create a new custom configuration file in a separate folder.
# 2. The first step in the new custom configuration file should be to load the default configuration file:
#    # Loading default replacement variables shipped with Set-OutlookSignatures
#    . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $(Join-Path -Path $(Get-Location).ProviderPath -ChildPath '\config\default replacement variables.ps1') -Raw)))
# 3. After importing the default configuration file, existing replacement variables can be altered with custom definitions and new replacement variables can be added.
# 4. Instead of altering existing replacement variables, it is recommended to create new replacement variables with modified content.
# 5. Start Set-OutlookSignatures with the parameter 'ReplacementVariableConfigFile' pointing to the new custom configuration file.


# Currently logged in user
$ReplaceHash['$CurrentUserGivenName$'] = [string]$ADPropsCurrentUser.givenName
$ReplaceHash['$CurrentUserSurname$'] = [string]$ADPropsCurrentUser.sn
$ReplaceHash['$CurrentUserDepartment$'] = [string]$ADPropsCurrentUser.department
$ReplaceHash['$CurrentUserTitle$'] = [string]$ADPropsCurrentUser.title
$ReplaceHash['$CurrentUserStreetAddress$'] = [string]$ADPropsCurrentUser.streetAddress
$ReplaceHash['$CurrentUserPostalcode$'] = [string]$ADPropsCurrentUser.postalCode
$ReplaceHash['$CurrentUserLocation$'] = [string]$ADPropsCurrentUser.l
$ReplaceHash['$CurrentUserCountry$'] = [string]$ADPropsCurrentUser.co
$ReplaceHash['$CurrentUserState$'] = [string]$ADPropsCurrentUser.st
$ReplaceHash['$CurrentUserTelephone$'] = [string]$ADPropsCurrentUser.telephoneNumber
$ReplaceHash['$CurrentUserFax$'] = [string]$ADPropsCurrentUser.facsimileTelephoneNumber
$ReplaceHash['$CurrentUserMobile$'] = [string]$ADPropsCurrentUser.mobile
$ReplaceHash['$CurrentUserMail$'] = [string]$ADPropsCurrentUser.mail
$ReplaceHash['$CurrentUserPhoto$'] = $ADPropsCurrentUser.thumbnailPhoto
$ReplaceHash['$CurrentUserPhotoDeleteEmpty$'] = $ADPropsCurrentUser.thumbnailPhoto
$ReplaceHash['$CurrentUserExtAttr1$'] = [string]$ADPropsCurrentUser.extensionAttribute1
$ReplaceHash['$CurrentUserExtAttr2$'] = [string]$ADPropsCurrentUser.extensionAttribute2
$ReplaceHash['$CurrentUserExtAttr3$'] = [string]$ADPropsCurrentUser.extensionAttribute3
$ReplaceHash['$CurrentUserExtAttr4$'] = [string]$ADPropsCurrentUser.extensionAttribute4
$ReplaceHash['$CurrentUserExtAttr5$'] = [string]$ADPropsCurrentUser.extensionAttribute5
$ReplaceHash['$CurrentUserExtAttr6$'] = [string]$ADPropsCurrentUser.extensionAttribute6
$ReplaceHash['$CurrentUserExtAttr7$'] = [string]$ADPropsCurrentUser.extensionAttribute7
$ReplaceHash['$CurrentUserExtAttr8$'] = [string]$ADPropsCurrentUser.extensionAttribute8
$ReplaceHash['$CurrentUserExtAttr9$'] = [string]$ADPropsCurrentUser.extensionAttribute9
$ReplaceHash['$CurrentUserExtAttr10$'] = [string]$ADPropsCurrentUser.extensionAttribute10
$ReplaceHash['$CurrentUserExtAttr11$'] = [string]$ADPropsCurrentUser.extensionAttribute11
$ReplaceHash['$CurrentUserExtAttr12$'] = [string]$ADPropsCurrentUser.extensionAttribute12
$ReplaceHash['$CurrentUserExtAttr13$'] = [string]$ADPropsCurrentUser.extensionAttribute13
$ReplaceHash['$CurrentUserExtAttr14$'] = [string]$ADPropsCurrentUser.extensionAttribute14
$ReplaceHash['$CurrentUserExtAttr15$'] = [string]$ADPropsCurrentUser.extensionAttribute15
$ReplaceHash['$CurrentUserOffice$'] = [string]$ADPropsCurrentUser.physicalDeliveryOfficeName
$ReplaceHash['$CurrentUserCompany$'] = [string]$ADPropsCurrentUser.company
$ReplaceHash['$CurrentUserMailNickname$'] = [string]$ADPropsCurrentUser.mailNickname
$ReplaceHash['$CurrentUserDisplayName$'] = [string]$ADPropsCurrentUser.displayName


# Manager of currently logged in user
$ReplaceHash['$CurrentUserManagerGivenName$'] = [string]$ADPropsCurrentUserManager.givenName
$ReplaceHash['$CurrentUserManagerSurname$'] = [string]$ADPropsCurrentUserManager.sn
$ReplaceHash['$CurrentUserManagerDepartment$'] = [string]$ADPropsCurrentUserManager.department
$ReplaceHash['$CurrentUserManagerTitle$'] = [string]$ADPropsCurrentUserManager.title
$ReplaceHash['$CurrentUserManagerStreetAddress$'] = [string]$ADPropsCurrentUserManager.streetAddress
$ReplaceHash['$CurrentUserManagerPostalcode$'] = [string]$ADPropsCurrentUserManager.postalCode
$ReplaceHash['$CurrentUserManagerLocation$'] = [string]$ADPropsCurrentUserManager.l
$ReplaceHash['$CurrentUserManagerCountry$'] = [string]$ADPropsCurrentUserManager.co
$ReplaceHash['$CurrentUserManagerState$'] = [string]$ADPropsCurrentUserManager.st
$ReplaceHash['$CurrentUserManagerTelephone$'] = [string]$ADPropsCurrentUserManager.telephoneNumber
$ReplaceHash['$CurrentUserManagerFax$'] = [string]$ADPropsCurrentUserManager.facsimileTelephoneNumber
$ReplaceHash['$CurrentUserManagerMobile$'] = [string]$ADPropsCurrentUserManager.mobile
$ReplaceHash['$CurrentUserManagerMail$'] = [string]$ADPropsCurrentUserManager.mail
$ReplaceHash['$CurrentUserManagerPhoto$'] = $ADPropsCurrentUserManager.thumbnailPhoto
$ReplaceHash['$CurrentUserManagerExtAttr1$'] = [string]$ADPropsCurrentUserManager.extensionAttribute1
$ReplaceHash['$CurrentUserManagerExtAttr2$'] = [string]$ADPropsCurrentUserManager.extensionAttribute2
$ReplaceHash['$CurrentUserManagerExtAttr3$'] = [string]$ADPropsCurrentUserManager.extensionAttribute3
$ReplaceHash['$CurrentUserManagerExtAttr4$'] = [string]$ADPropsCurrentUserManager.extensionAttribute4
$ReplaceHash['$CurrentUserManagerExtAttr5$'] = [string]$ADPropsCurrentUserManager.extensionAttribute5
$ReplaceHash['$CurrentUserManagerExtAttr6$'] = [string]$ADPropsCurrentUserManager.extensionAttribute6
$ReplaceHash['$CurrentUserManagerExtAttr7$'] = [string]$ADPropsCurrentUserManager.extensionAttribute7
$ReplaceHash['$CurrentUserManagerExtAttr8$'] = [string]$ADPropsCurrentUserManager.extensionAttribute8
$ReplaceHash['$CurrentUserManagerExtAttr9$'] = [string]$ADPropsCurrentUserManager.extensionAttribute9
$ReplaceHash['$CurrentUserManagerExtAttr10$'] = [string]$ADPropsCurrentUserManager.extensionAttribute10
$ReplaceHash['$CurrentUserManagerExtAttr11$'] = [string]$ADPropsCurrentUserManager.extensionAttribute11
$ReplaceHash['$CurrentUserManagerExtAttr12$'] = [string]$ADPropsCurrentUserManager.extensionAttribute12
$ReplaceHash['$CurrentUserManagerExtAttr13$'] = [string]$ADPropsCurrentUserManager.extensionAttribute13
$ReplaceHash['$CurrentUserManagerExtAttr14$'] = [string]$ADPropsCurrentUserManager.extensionAttribute14
$ReplaceHash['$CurrentUserManagerExtAttr15$'] = [string]$ADPropsCurrentUserManager.extensionAttribute15
$ReplaceHash['$CurrentUserManagerOffice$'] = [string]$ADPropsCurrentUserManager.physicalDeliveryOfficeName
$ReplaceHash['$CurrentUserManagerCompany$'] = [string]$ADPropsCurrentUserManager.company
$ReplaceHash['$CurrentUserManagerMailNickname$'] = [string]$ADPropsCurrentUserManager.mailNickname
$ReplaceHash['$CurrentUserManagerDisplayName$'] = [string]$ADPropsCurrentUserManager.displayName


# Current mailbox
$ReplaceHash['$CurrentMailboxGivenName$'] = [string]$ADPropsCurrentMailbox.givenName
$ReplaceHash['$CurrentMailboxSurname$'] = [string]$ADPropsCurrentMailbox.sn
$ReplaceHash['$CurrentMailboxDepartment$'] = [string]$ADPropsCurrentMailbox.department
$ReplaceHash['$CurrentMailboxTitle$'] = [string]$ADPropsCurrentMailbox.title
$ReplaceHash['$CurrentMailboxStreetAddress$'] = [string]$ADPropsCurrentMailbox.streetAddress
$ReplaceHash['$CurrentMailboxPostalcode$'] = [string]$ADPropsCurrentMailbox.postalCode
$ReplaceHash['$CurrentMailboxLocation$'] = [string]$ADPropsCurrentMailbox.l
$ReplaceHash['$CurrentMailboxCountry$'] = [string]$ADPropsCurrentMailbox.co
$ReplaceHash['$CurrentMailboxState$'] = [string]$ADPropsCurrentMailbox.st
$ReplaceHash['$CurrentMailboxTelephone$'] = [string]$ADPropsCurrentMailbox.telephoneNumber
$ReplaceHash['$CurrentMailboxFax$'] = [string]$ADPropsCurrentMailbox.facsimileTelephoneNumber
$ReplaceHash['$CurrentMailboxMobile$'] = [string]$ADPropsCurrentMailbox.mobile
$ReplaceHash['$CurrentMailboxMail$'] = [string]$ADPropsCurrentMailbox.mail
$ReplaceHash['$CurrentMailboxPhoto$'] = $ADPropsCurrentMailbox.thumbnailPhoto
$ReplaceHash['$CurrentMailboxExtAttr1$'] = [string]$ADPropsCurrentMailbox.extensionAttribute1
$ReplaceHash['$CurrentMailboxExtAttr2$'] = [string]$ADPropsCurrentMailbox.extensionAttribute2
$ReplaceHash['$CurrentMailboxExtAttr3$'] = [string]$ADPropsCurrentMailbox.extensionAttribute3
$ReplaceHash['$CurrentMailboxExtAttr4$'] = [string]$ADPropsCurrentMailbox.extensionAttribute4
$ReplaceHash['$CurrentMailboxExtAttr5$'] = [string]$ADPropsCurrentMailbox.extensionAttribute5
$ReplaceHash['$CurrentMailboxExtAttr6$'] = [string]$ADPropsCurrentMailbox.extensionAttribute6
$ReplaceHash['$CurrentMailboxExtAttr7$'] = [string]$ADPropsCurrentMailbox.extensionAttribute7
$ReplaceHash['$CurrentMailboxExtAttr8$'] = [string]$ADPropsCurrentMailbox.extensionAttribute8
$ReplaceHash['$CurrentMailboxExtAttr9$'] = [string]$ADPropsCurrentMailbox.extensionAttribute9
$ReplaceHash['$CurrentMailboxExtAttr10$'] = [string]$ADPropsCurrentMailbox.extensionAttribute10
$ReplaceHash['$CurrentMailboxExtAttr11$'] = [string]$ADPropsCurrentMailbox.extensionAttribute11
$ReplaceHash['$CurrentMailboxExtAttr12$'] = [string]$ADPropsCurrentMailbox.extensionAttribute12
$ReplaceHash['$CurrentMailboxExtAttr13$'] = [string]$ADPropsCurrentMailbox.extensionAttribute13
$ReplaceHash['$CurrentMailboxExtAttr14$'] = [string]$ADPropsCurrentMailbox.extensionAttribute14
$ReplaceHash['$CurrentMailboxExtAttr15$'] = [string]$ADPropsCurrentMailbox.extensionAttribute15
$ReplaceHash['$CurrentMailboxOffice$'] = [string]$ADPropsCurrentMailbox.physicalDeliveryOfficeName
$ReplaceHash['$CurrentMailboxCompany$'] = [string]$ADPropsCurrentMailbox.company
$ReplaceHash['$CurrentMailboxMailNickname$'] = [string]$ADPropsCurrentMailbox.mailNickname
$ReplaceHash['$CurrentMailboxDisplayName$'] = [string]$ADPropsCurrentMailbox.displayName


# Manager of current mailbox
$ReplaceHash['$CurrentMailboxManagerGivenName$'] = [string]$ADPropsCurrentMailboxManager.givenName
$ReplaceHash['$CurrentMailboxManagerSurname$'] = [string]$ADPropsCurrentMailboxManager.sn
$ReplaceHash['$CurrentMailboxManagerDepartment$'] = [string]$ADPropsCurrentMailboxManager.department
$ReplaceHash['$CurrentMailboxManagerTitle$'] = [string]$ADPropsCurrentMailboxManager.title
$ReplaceHash['$CurrentMailboxManagerStreetAddress$'] = [string]$ADPropsCurrentMailboxManager.streetAddress
$ReplaceHash['$CurrentMailboxManagerPostalcode$'] = [string]$ADPropsCurrentMailboxManager.postalCode
$ReplaceHash['$CurrentMailboxManagerLocation$'] = [string]$ADPropsCurrentMailboxManager.l
$ReplaceHash['$CurrentMailboxManagerCountry$'] = [string]$ADPropsCurrentMailboxManager.co
$ReplaceHash['$CurrentMailboxManagerState$'] = [string]$ADPropsCurrentMailboxManager.st
$ReplaceHash['$CurrentMailboxManagerTelephone$'] = [string]$ADPropsCurrentMailboxManager.telephoneNumber
$ReplaceHash['$CurrentMailboxManagerFax$'] = [string]$ADPropsCurrentMailboxManager.facsimileTelephoneNumber
$ReplaceHash['$CurrentMailboxManagerMobile$'] = [string]$ADPropsCurrentMailboxManager.mobile
$ReplaceHash['$CurrentMailboxManagerMail$'] = [string]$ADPropsCurrentMailboxManager.mail
$ReplaceHash['$CurrentMailboxManagerPhoto$'] = $ADPropsCurrentMailboxManager.thumbnailPhoto
$ReplaceHash['$CurrentMailboxManagerExtAttr1$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute1
$ReplaceHash['$CurrentMailboxManagerExtAttr2$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute2
$ReplaceHash['$CurrentMailboxManagerExtAttr3$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute3
$ReplaceHash['$CurrentMailboxManagerExtAttr4$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute4
$ReplaceHash['$CurrentMailboxManagerExtAttr5$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute5
$ReplaceHash['$CurrentMailboxManagerExtAttr6$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute6
$ReplaceHash['$CurrentMailboxManagerExtAttr7$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute7
$ReplaceHash['$CurrentMailboxManagerExtAttr8$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute8
$ReplaceHash['$CurrentMailboxManagerExtAttr9$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute9
$ReplaceHash['$CurrentMailboxManagerExtAttr10$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute10
$ReplaceHash['$CurrentMailboxManagerExtAttr11$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute11
$ReplaceHash['$CurrentMailboxManagerExtAttr12$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute12
$ReplaceHash['$CurrentMailboxManagerExtAttr13$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute13
$ReplaceHash['$CurrentMailboxManagerExtAttr14$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute14
$ReplaceHash['$CurrentMailboxManagerExtAttr15$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute15
$ReplaceHash['$CurrentMailboxManagerOffice$'] = [string]$ADPropsCurrentMailboxManager.physicalDeliveryOfficeName
$ReplaceHash['$CurrentMailboxManagerCompany$'] = [string]$ADPropsCurrentMailboxManager.company
$ReplaceHash['$CurrentMailboxManagerMailNickname$'] = [string]$ADPropsCurrentMailboxManager.mailNickname
$ReplaceHash['$CurrentMailboxManagerDisplayName$'] = [string]$ADPropsCurrentMailboxManager.displayName


# Sample code: Full user name including honorific and academic titles
#   $CurrentUserNameWithHonorifics$, $CurrentUserManagerNameWithHonorifics$, $CurrentMailboxNameWithHonorifics$, $CurrentMailboxManagerNameWithHonorifics$
# According to standards in German speaking countries:
#   "<custom AD attribute 'honorificPrefix'> <standard AD attribute 'givenname'> <standard AD attribute 'surname'>, <custom AD attribute 'honorificSuffix'>"
# If one or more attributes are not set, unnecessary whitespaces and commas are avoided
# Examples:
#   Mag. Dr. John Doe, BA MA PhD
#   Dr. John Doe
#   John Doe, PhD
#   John Doe
# Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers commercial support for this and other open source code.
$ReplaceHash['$CurrentUserNameWithHonorifics$'] = (((((([string]$ADPropsCurrentUser.honorificPrefix, [string]$ADPropsCurrentUser.givenname, [string]$ADPropsCurrentUser.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUser.honorificSuffix) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CurrentUserManagerNameWithHonorifics$'] = (((((([string]$ADPropsCurrentUserManager.honorificPrefix, [string]$ADPropsCurrentUserManager.givenname, [string]$ADPropsCurrentUserManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUserManager.honorificSuffix) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CurrentMailboxNameWithHonorifics$'] = (((((([string]$ADPropsCurrentMailbox.honorificPrefix, [string]$ADPropsCurrentMailbox.givenname, [string]$ADPropsCurrentMailbox.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailbox.honorificSuffix) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CurrentMailboxManagerNameWithHonorifics$'] = (((((([string]$ADPropsCurrentMailboxManager.honorificPrefix, [string]$ADPropsCurrentMailboxManager.givenname, [string]$ADPropsCurrentMailboxManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailboxManager.honorificSuffix) | Where-Object { $_ -ne '' }) -join ', ')


# Sample code: Create MeCard (vCard alternative) QR codes and save the images in the following replacement variables:
#   $CurrentUserCustomImage1$, $CurrentUserManagerCustomImage1$, $CurrentMailboxCustomImage1$, $CurrentMailboxManagerCustomImage1$
# Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers commercial support for this and other open source code.
@('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager') | ForEach-Object {
    $QRCodeContent = @(
        @(
            @(
                'MECARD:'
                "N:$($ReplaceHash['$' + $_ + 'Surname$']),$($ReplaceHash['$' + $_ + 'GivenName$']);"
                "NOTE:$($ReplaceHash['$' + $_ + 'Company$'])"
                "$($ReplaceHash['$' + $_ + 'Title$']);"
                "EMAIL:$($ReplaceHash['$' + $_ + 'Mail$']);"
                "TEL:$($ReplaceHash['$' + $_ + 'Mobile$']);"
                "ADR:$($ReplaceHash['$' + $_ + 'StreetAddress$'])"
                "$("$($ReplaceHash['$' + $_ + 'Postalcode$']) $($ReplaceHash['$' + $_ + 'Location$'])")"
                "$($ReplaceHash['$' + $_ + 'State$'])"
                "$($ReplaceHash['$' + $_ + 'Country$']);"
                'URL:explicitconsulting.at;'
                ';'
            ) | ForEach-Object { $_.trim() }
        ) | Where-Object { $_ -and (-not $_.EndsWith(':;')) -and (-not $_.EndsWith(':,;')) }
    ) -join ("`r`n") -replace ':\r\n;', ':;' -replace '\r\n(.*):;', ''

    if ($QRCodeContent -notmatch '\r\nN:.*;\r\n') { $QRCodeContent = 'https://explicitconsulting.at' }

    $ReplaceHash['$' + $_ + 'CustomImage1$'] = ((New-Object -TypeName QRCoder.PngByteQRCode -ArgumentList ((New-Object -TypeName QRCoder.QRCodeGenerator).CreateQrCode($QRCodeContent, 'L', $true))).GetGraphic(20, [byte[]]@(0, 0, 0), [byte[]]@(255, 255, 255), $false))
}


# Sample code: Take salutation or gender pronouns string from Extension Attribute 3
#   $CurrentUserSalutation$, $CurrentUserManagerSalutation$, $CurrentMailboxSalutation$, $CurrentMailboxManagerSalutation$
#   $CurrentUserGenderPronouns$, $CurrentUserManagerGenderPronouns$, $CurrentMailboxGenderPronouns$, $CurrentMailboxManagerGenderPronouns$
# Format
#   If ExtensionAttribute3 is not empty or whitespace, put it in brackets and add a leading space
#     Examples: " (Mr.)", " (Ms.)", " (she/her)"
#   Else: '' (emtpy string)
# Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers commercial support for this and other open source code.
$ReplaceHash['$CurrentUserSalutation$'] = $ReplaceHash['$CurrentUserGenderPronouns$'] = $(if ([string]::IsNullOrWhiteSpace([string]$ADPropsCurrentUser.extensionattribute3)) { $null } else { " ($([string]$ADPropsCurrentUser.extensionattribute3))" })
$ReplaceHash['$CurrentUserManagerSalutation$'] = $ReplaceHash['$CurrentUserManagerGenderPronouns$'] = $(if ([string]::IsNullOrWhiteSpace([string]$ADPropsCurrentUserManager.extensionattribute3)) { $null } else { " ($([string]$ADPropsCurrentUserManager.extensionattribute3))" })
$ReplaceHash['$CurrentMailboxSalutation$'] = $ReplaceHash['$CurrentMailboxGenderPronouns$'] = $(if ([string]::IsNullOrWhiteSpace([string]$ADPropsCurrentMailbox.extensionattribute3)) { $null } else { " ($([string]$ADPropsCurrentMailbox.extensionattribute3))" })
$ReplaceHash['$CurrentMailboxManagerSalutation$'] = $ReplaceHash['$CurrentMailboxManagerGenderPronouns$'] = $(if ([string]::IsNullOrWhiteSpace([string]$ADPropsCurrentMailboxManager.extensionattribute3)) { $null } else { " ($([string]$ADPropsCurrentMailboxManager.extensionattribute3))" })
