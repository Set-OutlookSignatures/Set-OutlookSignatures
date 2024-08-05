# This file allows defining custom replacement variables for Set-OutlookSignatures
#
# This script is executed as a whole once for each mailbox.
# It allows for complex replacement variable handling (complex string transformations, retrieving information from web services and databases, etc.).
#
# Attention: The configuration file is executed as part of Set-OutlookSignatures.ps1 and is not checked for any harmful content. Please only allow qualified technicians write access to this file, only use it to define replacement variables and test it thoroughly.
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
$ReplaceHash['$CurrentUserGivenname$'] = [string]$ADPropsCurrentUser.givenName
$ReplaceHash['$CurrentUserSurname$'] = [string]$ADPropsCurrentUser.sn
$ReplaceHash['$CurrentUserDepartment$'] = [string]$ADPropsCurrentUser.department
$ReplaceHash['$CurrentUserTitle$'] = [string]$ADPropsCurrentUser.title
$ReplaceHash['$CurrentUserStreetaddress$'] = [string]$ADPropsCurrentUser.streetAddress
$ReplaceHash['$CurrentUserPostalcode$'] = [string]$ADPropsCurrentUser.postalCode
$ReplaceHash['$CurrentUserLocation$'] = [string]$ADPropsCurrentUser.l
$ReplaceHash['$CurrentUserCountry$'] = [string]$ADPropsCurrentUser.co
$ReplaceHash['$CurrentUserState$'] = [string]$ADPropsCurrentUser.st
$ReplaceHash['$CurrentUserTelephone$'] = [string]$ADPropsCurrentUser.telephoneNumber
$ReplaceHash['$CurrentUserFax$'] = [string]$ADPropsCurrentUser.facsimileTelephoneNumber
$ReplaceHash['$CurrentUserMobile$'] = [string]$ADPropsCurrentUser.mobile
$ReplaceHash['$CurrentUserMail$'] = [string]$ADPropsCurrentUser.mail
$ReplaceHash['$CurrentUserPhoto$'] = $ADPropsCurrentUser.thumbnailPhoto
$ReplaceHash['$CurrentUserPhotodeleteempty$'] = $ADPropsCurrentUser.thumbnailPhoto
$ReplaceHash['$CurrentUserExtattr1$'] = [string]$ADPropsCurrentUser.extensionAttribute1
$ReplaceHash['$CurrentUserExtattr2$'] = [string]$ADPropsCurrentUser.extensionAttribute2
$ReplaceHash['$CurrentUserExtattr3$'] = [string]$ADPropsCurrentUser.extensionAttribute3
$ReplaceHash['$CurrentUserExtattr4$'] = [string]$ADPropsCurrentUser.extensionAttribute4
$ReplaceHash['$CurrentUserExtattr5$'] = [string]$ADPropsCurrentUser.extensionAttribute5
$ReplaceHash['$CurrentUserExtattr6$'] = [string]$ADPropsCurrentUser.extensionAttribute6
$ReplaceHash['$CurrentUserExtattr7$'] = [string]$ADPropsCurrentUser.extensionAttribute7
$ReplaceHash['$CurrentUserExtattr8$'] = [string]$ADPropsCurrentUser.extensionAttribute8
$ReplaceHash['$CurrentUserExtattr9$'] = [string]$ADPropsCurrentUser.extensionAttribute9
$ReplaceHash['$CurrentUserExtattr10$'] = [string]$ADPropsCurrentUser.extensionAttribute10
$ReplaceHash['$CurrentUserExtattr11$'] = [string]$ADPropsCurrentUser.extensionAttribute11
$ReplaceHash['$CurrentUserExtattr12$'] = [string]$ADPropsCurrentUser.extensionAttribute12
$ReplaceHash['$CurrentUserExtattr13$'] = [string]$ADPropsCurrentUser.extensionAttribute13
$ReplaceHash['$CurrentUserExtattr14$'] = [string]$ADPropsCurrentUser.extensionAttribute14
$ReplaceHash['$CurrentUserExtattr15$'] = [string]$ADPropsCurrentUser.extensionAttribute15
$ReplaceHash['$CurrentUserOffice$'] = [string]$ADPropsCurrentUser.physicalDeliveryOfficeName
$ReplaceHash['$CurrentUserCompany$'] = [string]$ADPropsCurrentUser.company
$ReplaceHash['$CurrentUserMailnickname$'] = [string]$ADPropsCurrentUser.mailNickname
$ReplaceHash['$CurrentUserDisplayname$'] = [string]$ADPropsCurrentUser.displayName


# Manager of currently logged in user
$ReplaceHash['$CurrentUserManagerGivenname$'] = [string]$ADPropsCurrentUserManager.givenName
$ReplaceHash['$CurrentUserManagerSurname$'] = [string]$ADPropsCurrentUserManager.sn
$ReplaceHash['$CurrentUserManagerDepartment$'] = [string]$ADPropsCurrentUserManager.department
$ReplaceHash['$CurrentUserManagerTitle$'] = [string]$ADPropsCurrentUserManager.title
$ReplaceHash['$CurrentUserManagerStreetaddress$'] = [string]$ADPropsCurrentUserManager.streetAddress
$ReplaceHash['$CurrentUserManagerPostalcode$'] = [string]$ADPropsCurrentUserManager.postalCode
$ReplaceHash['$CurrentUserManagerLocation$'] = [string]$ADPropsCurrentUserManager.l
$ReplaceHash['$CurrentUserManagerCountry$'] = [string]$ADPropsCurrentUserManager.co
$ReplaceHash['$CurrentUserManagerState$'] = [string]$ADPropsCurrentUserManager.st
$ReplaceHash['$CurrentUserManagerTelephone$'] = [string]$ADPropsCurrentUserManager.telephoneNumber
$ReplaceHash['$CurrentUserManagerFax$'] = [string]$ADPropsCurrentUserManager.facsimileTelephoneNumber
$ReplaceHash['$CurrentUserManagerMobile$'] = [string]$ADPropsCurrentUserManager.mobile
$ReplaceHash['$CurrentUserManagerMail$'] = [string]$ADPropsCurrentUserManager.mail
$ReplaceHash['$CurrentUserManagerPhoto$'] = $ADPropsCurrentUserManager.thumbnailPhoto
$ReplaceHash['$CurrentUserManagerExtattr1$'] = [string]$ADPropsCurrentUserManager.extensionAttribute1
$ReplaceHash['$CurrentUserManagerExtattr2$'] = [string]$ADPropsCurrentUserManager.extensionAttribute2
$ReplaceHash['$CurrentUserManagerExtattr3$'] = [string]$ADPropsCurrentUserManager.extensionAttribute3
$ReplaceHash['$CurrentUserManagerExtattr4$'] = [string]$ADPropsCurrentUserManager.extensionAttribute4
$ReplaceHash['$CurrentUserManagerExtattr5$'] = [string]$ADPropsCurrentUserManager.extensionAttribute5
$ReplaceHash['$CurrentUserManagerExtattr6$'] = [string]$ADPropsCurrentUserManager.extensionAttribute6
$ReplaceHash['$CurrentUserManagerExtattr7$'] = [string]$ADPropsCurrentUserManager.extensionAttribute7
$ReplaceHash['$CurrentUserManagerExtattr8$'] = [string]$ADPropsCurrentUserManager.extensionAttribute8
$ReplaceHash['$CurrentUserManagerExtattr9$'] = [string]$ADPropsCurrentUserManager.extensionAttribute9
$ReplaceHash['$CurrentUserManagerExtattr10$'] = [string]$ADPropsCurrentUserManager.extensionAttribute10
$ReplaceHash['$CurrentUserManagerExtattr11$'] = [string]$ADPropsCurrentUserManager.extensionAttribute11
$ReplaceHash['$CurrentUserManagerExtattr12$'] = [string]$ADPropsCurrentUserManager.extensionAttribute12
$ReplaceHash['$CurrentUserManagerExtattr13$'] = [string]$ADPropsCurrentUserManager.extensionAttribute13
$ReplaceHash['$CurrentUserManagerExtattr14$'] = [string]$ADPropsCurrentUserManager.extensionAttribute14
$ReplaceHash['$CurrentUserManagerExtattr15$'] = [string]$ADPropsCurrentUserManager.extensionAttribute15
$ReplaceHash['$CurrentUserManagerOffice$'] = [string]$ADPropsCurrentUserManager.physicalDeliveryOfficeName
$ReplaceHash['$CurrentUserManagerCompany$'] = [string]$ADPropsCurrentUserManager.company
$ReplaceHash['$CurrentUserManagerMailnickname$'] = [string]$ADPropsCurrentUserManager.mailNickname
$ReplaceHash['$CurrentUserManagerDisplayname$'] = [string]$ADPropsCurrentUserManager.displayName


# Current mailbox
$ReplaceHash['$CurrentMailboxGivenname$'] = [string]$ADPropsCurrentMailbox.givenName
$ReplaceHash['$CurrentMailboxSurname$'] = [string]$ADPropsCurrentMailbox.sn
$ReplaceHash['$CurrentMailboxDepartment$'] = [string]$ADPropsCurrentMailbox.department
$ReplaceHash['$CurrentMailboxTitle$'] = [string]$ADPropsCurrentMailbox.title
$ReplaceHash['$CurrentMailboxStreetaddress$'] = [string]$ADPropsCurrentMailbox.streetAddress
$ReplaceHash['$CurrentMailboxPostalcode$'] = [string]$ADPropsCurrentMailbox.postalCode
$ReplaceHash['$CurrentMailboxLocation$'] = [string]$ADPropsCurrentMailbox.l
$ReplaceHash['$CurrentMailboxCountry$'] = [string]$ADPropsCurrentMailbox.co
$ReplaceHash['$CurrentMailboxState$'] = [string]$ADPropsCurrentMailbox.st
$ReplaceHash['$CurrentMailboxTelephone$'] = [string]$ADPropsCurrentMailbox.telephoneNumber
$ReplaceHash['$CurrentMailboxFax$'] = [string]$ADPropsCurrentMailbox.facsimileTelephoneNumber
$ReplaceHash['$CurrentMailboxMobile$'] = [string]$ADPropsCurrentMailbox.mobile
$ReplaceHash['$CurrentMailboxMail$'] = [string]$ADPropsCurrentMailbox.mail
$ReplaceHash['$CurrentMailboxPhoto$'] = $ADPropsCurrentMailbox.thumbnailPhoto
$ReplaceHash['$CurrentMailboxExtattr1$'] = [string]$ADPropsCurrentMailbox.extensionAttribute1
$ReplaceHash['$CurrentMailboxExtattr2$'] = [string]$ADPropsCurrentMailbox.extensionAttribute2
$ReplaceHash['$CurrentMailboxExtattr3$'] = [string]$ADPropsCurrentMailbox.extensionAttribute3
$ReplaceHash['$CurrentMailboxExtattr4$'] = [string]$ADPropsCurrentMailbox.extensionAttribute4
$ReplaceHash['$CurrentMailboxExtattr5$'] = [string]$ADPropsCurrentMailbox.extensionAttribute5
$ReplaceHash['$CurrentMailboxExtattr6$'] = [string]$ADPropsCurrentMailbox.extensionAttribute6
$ReplaceHash['$CurrentMailboxExtattr7$'] = [string]$ADPropsCurrentMailbox.extensionAttribute7
$ReplaceHash['$CurrentMailboxExtattr8$'] = [string]$ADPropsCurrentMailbox.extensionAttribute8
$ReplaceHash['$CurrentMailboxExtattr9$'] = [string]$ADPropsCurrentMailbox.extensionAttribute9
$ReplaceHash['$CurrentMailboxExtattr10$'] = [string]$ADPropsCurrentMailbox.extensionAttribute10
$ReplaceHash['$CurrentMailboxExtattr11$'] = [string]$ADPropsCurrentMailbox.extensionAttribute11
$ReplaceHash['$CurrentMailboxExtattr12$'] = [string]$ADPropsCurrentMailbox.extensionAttribute12
$ReplaceHash['$CurrentMailboxExtattr13$'] = [string]$ADPropsCurrentMailbox.extensionAttribute13
$ReplaceHash['$CurrentMailboxExtattr14$'] = [string]$ADPropsCurrentMailbox.extensionAttribute14
$ReplaceHash['$CurrentMailboxExtattr15$'] = [string]$ADPropsCurrentMailbox.extensionAttribute15
$ReplaceHash['$CurrentMailboxOffice$'] = [string]$ADPropsCurrentMailbox.physicalDeliveryOfficeName
$ReplaceHash['$CurrentMailboxCompany$'] = [string]$ADPropsCurrentMailbox.company
$ReplaceHash['$CurrentMailboxMailnickname$'] = [string]$ADPropsCurrentMailbox.mailNickname
$ReplaceHash['$CurrentMailboxDisplayname$'] = [string]$ADPropsCurrentMailbox.displayName


# Manager of current mailbox
$ReplaceHash['$CurrentMailboxManagerGivenname$'] = [string]$ADPropsCurrentMailboxManager.givenName
$ReplaceHash['$CurrentMailboxManagerSurname$'] = [string]$ADPropsCurrentMailboxManager.sn
$ReplaceHash['$CurrentMailboxManagerDepartment$'] = [string]$ADPropsCurrentMailboxManager.department
$ReplaceHash['$CurrentMailboxManagerTitle$'] = [string]$ADPropsCurrentMailboxManager.title
$ReplaceHash['$CurrentMailboxManagerStreetaddress$'] = [string]$ADPropsCurrentMailboxManager.streetAddress
$ReplaceHash['$CurrentMailboxManagerPostalcode$'] = [string]$ADPropsCurrentMailboxManager.postalCode
$ReplaceHash['$CurrentMailboxManagerLocation$'] = [string]$ADPropsCurrentMailboxManager.l
$ReplaceHash['$CurrentMailboxManagerCountry$'] = [string]$ADPropsCurrentMailboxManager.co
$ReplaceHash['$CurrentMailboxManagerState$'] = [string]$ADPropsCurrentMailboxManager.st
$ReplaceHash['$CurrentMailboxManagerTelephone$'] = [string]$ADPropsCurrentMailboxManager.telephoneNumber
$ReplaceHash['$CurrentMailboxManagerFax$'] = [string]$ADPropsCurrentMailboxManager.facsimileTelephoneNumber
$ReplaceHash['$CurrentMailboxManagerMobile$'] = [string]$ADPropsCurrentMailboxManager.mobile
$ReplaceHash['$CurrentMailboxManagerMail$'] = [string]$ADPropsCurrentMailboxManager.mail
$ReplaceHash['$CurrentMailboxManagerPhoto$'] = $ADPropsCurrentMailboxManager.thumbnailPhoto
$ReplaceHash['$CurrentMailboxManagerExtattr1$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute1
$ReplaceHash['$CurrentMailboxManagerExtattr2$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute2
$ReplaceHash['$CurrentMailboxManagerExtattr3$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute3
$ReplaceHash['$CurrentMailboxManagerExtattr4$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute4
$ReplaceHash['$CurrentMailboxManagerExtattr5$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute5
$ReplaceHash['$CurrentMailboxManagerExtattr6$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute6
$ReplaceHash['$CurrentMailboxManagerExtattr7$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute7
$ReplaceHash['$CurrentMailboxManagerExtattr8$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute8
$ReplaceHash['$CurrentMailboxManagerExtattr9$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute9
$ReplaceHash['$CurrentMailboxManagerExtattr10$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute10
$ReplaceHash['$CurrentMailboxManagerExtattr11$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute11
$ReplaceHash['$CurrentMailboxManagerExtattr12$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute12
$ReplaceHash['$CurrentMailboxManagerExtattr13$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute13
$ReplaceHash['$CurrentMailboxManagerExtattr14$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute14
$ReplaceHash['$CurrentMailboxManagerExtattr15$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute15
$ReplaceHash['$CurrentMailboxManagerOffice$'] = [string]$ADPropsCurrentMailboxManager.physicalDeliveryOfficeName
$ReplaceHash['$CurrentMailboxManagerCompany$'] = [string]$ADPropsCurrentMailboxManager.company
$ReplaceHash['$CurrentMailboxManagerMailnickname$'] = [string]$ADPropsCurrentMailboxManager.mailNickname
$ReplaceHash['$CurrentMailboxManagerDisplayname$'] = [string]$ADPropsCurrentMailboxManager.displayName


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
# Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers fee-based support for this and other open source code.
$ReplaceHash['$CurrentUserNameWithHonorifics$'] = (((((([string]$ADPropsCurrentUser.honorificPrefix, [string]$ADPropsCurrentUser.givenname, [string]$ADPropsCurrentUser.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUser.honorificSuffix) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CurrentUserManagerNameWithHonorifics$'] = (((((([string]$ADPropsCurrentUserManager.honorificPrefix, [string]$ADPropsCurrentUserManager.givenname, [string]$ADPropsCurrentUserManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUserManager.honorificSuffix) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CurrentMailboxNameWithHonorifics$'] = (((((([string]$ADPropsCurrentMailbox.honorificPrefix, [string]$ADPropsCurrentMailbox.givenname, [string]$ADPropsCurrentMailbox.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailbox.honorificSuffix) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CurrentMailboxManagerNameWithHonorifics$'] = (((((([string]$ADPropsCurrentMailboxManager.honorificPrefix, [string]$ADPropsCurrentMailboxManager.givenname, [string]$ADPropsCurrentMailboxManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailboxManager.honorificSuffix) | Where-Object { $_ -ne '' }) -join ', ')


# Sample code: Create MeCard (vCard alternative) QR codes and save the images in the following replacement variables:
#   $CurrentUserCustomImage1$, $CurrentUserManagerCustomImage1$, $CurrentMailboxCustomImage1$, $CurrentMailboxManagerCustomImage1$
# Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers fee-based support for this and other open source code.
@('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager') | ForEach-Object {
    $QRCodeContent = @(
        @(
            @(
                'MECARD:'
                "N:$($ReplaceHash['$' + $_ + 'Surname$']),$($ReplaceHash['$' + $_ + 'Givenname$']);"
                "NOTE:$($ReplaceHash['$' + $_ + 'Company$'])"
                "$($ReplaceHash['$' + $_ + 'Title$']);"
                "EMAIL:$($ReplaceHash['$' + $_ + 'Mail$']);"
                "TEL:$($ReplaceHash['$' + $_ + 'Mobile$']);"
                "ADR:$($ReplaceHash['$' + $_ + 'Streetaddress$'])"
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


# Sample code: Create gender pronouns string from Extension Attribute 3
#   $CurrentUserGenderPronouns$, $CurrentUserManagerGenderPronouns$, # $CurrentMailboxGenderPronouns$, $CurrentMailboxManagerGenderPronouns$
# Format
#   ExtensionAttribute3 contains at least three characters, and a forward slash somewhere between the first and last character: " (<ExtensionAttribute3>)"
#     Examples: " (she/her)", " (he/him)"
#   Else: '' (emtpy string)
# Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers fee-based support for this and other open source code.
$ReplaceHash['$CurrentUserGenderPronouns$'] = $(if (([string]$ADPropsCurrentUser.ExtensionAttribute3) -imatch '.+\/.+') { " ($(([string]$ADPropsCurrentUser.ExtensionAttribute3)))" } else { '' })
$ReplaceHash['$CurrentUserManagerGenderPronouns$'] = $(if (([string]$ADPropsCurrentUserManager.ExtensionAttribute3) -imatch '.+\/.+') { " ($(([string]$ADPropsCurrentUserManager.ExtensionAttribute3)))" } else { '' })
$ReplaceHash['$CurrentMailboxGenderPronouns$'] = $(if (([string]$ADPropsCurrentMailbox.ExtensionAttribute3) -imatch '.+\/.+') { " ($(([string]$ADPropsCurrentMailbox.ExtensionAttribute3)))" } else { '' })
$ReplaceHash['$CurrentMailboxManagerGenderPronouns$'] = $(if (([string]$ADPropsCurrentMailboxManager.ExtensionAttribute3) -imatch '.+\/.+') { " ($(([string]$ADPropsCurrentMailboxManager.ExtensionAttribute3)))" } else { '' })
