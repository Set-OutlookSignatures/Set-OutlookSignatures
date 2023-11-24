# This file allows defining custom replacement variables for Set-OutlookSignatures
#
# This script is executed as a whole once for each mailbox.
# It allows for complex replacement variable handling (complex string transformations, retrieving information from web services and databases, etc.).
#
# Attention: The configuration file is executed as part of Set-OutlookSignatures.ps1 and is not checked for any harmful content. Please only allow qualified technicians write access to this file, only use it to to define replacement variables and test it thoroughly.
#
# Replacement variable names are not case sensitive.
#
# Active Directory property names are case sensitive.
# It is required to use full lowercase Active Directory property names.
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
$ReplaceHash['$CurrentUserGivenname$'] = [string]$ADPropsCurrentUser.givenname
$ReplaceHash['$CurrentUserSurname$'] = [string]$ADPropsCurrentUser.sn
$ReplaceHash['$CurrentUserDepartment$'] = [string]$ADPropsCurrentUser.department
$ReplaceHash['$CurrentUserTitle$'] = [string]$ADPropsCurrentUser.title
$ReplaceHash['$CurrentUserStreetaddress$'] = [string]$ADPropsCurrentUser.streetaddress
$ReplaceHash['$CurrentUserPostalcode$'] = [string]$ADPropsCurrentUser.postalcode
$ReplaceHash['$CurrentUserLocation$'] = [string]$ADPropsCurrentUser.l
$ReplaceHash['$CurrentUserCountry$'] = [string]$ADPropsCurrentUser.co
$ReplaceHash['$CurrentUserState$'] = [string]$ADPropsCurrentUser.st
$ReplaceHash['$CurrentUserTelephone$'] = [string]$ADPropsCurrentUser.telephonenumber
$ReplaceHash['$CurrentUserFax$'] = [string]$ADPropsCurrentUser.facsimiletelephonenumber
$ReplaceHash['$CurrentUserMobile$'] = [string]$ADPropsCurrentUser.mobile
$ReplaceHash['$CurrentUserMail$'] = [string]$ADPropsCurrentUser.mail
$ReplaceHash['$CurrentUserPhoto$'] = $ADPropsCurrentUser.thumbnailphoto
$ReplaceHash['$CurrentUserPhotodeleteempty$'] = $ADPropsCurrentUser.thumbnailphoto
$ReplaceHash['$CurrentUserExtattr1$'] = [string]$ADPropsCurrentUser.extensionattribute1
$ReplaceHash['$CurrentUserExtattr2$'] = [string]$ADPropsCurrentUser.extensionattribute2
$ReplaceHash['$CurrentUserExtattr3$'] = [string]$ADPropsCurrentUser.extensionattribute3
$ReplaceHash['$CurrentUserExtattr4$'] = [string]$ADPropsCurrentUser.extensionattribute4
$ReplaceHash['$CurrentUserExtattr5$'] = [string]$ADPropsCurrentUser.extensionattribute5
$ReplaceHash['$CurrentUserExtattr6$'] = [string]$ADPropsCurrentUser.extensionattribute6
$ReplaceHash['$CurrentUserExtattr7$'] = [string]$ADPropsCurrentUser.extensionattribute7
$ReplaceHash['$CurrentUserExtattr8$'] = [string]$ADPropsCurrentUser.extensionattribute8
$ReplaceHash['$CurrentUserExtattr9$'] = [string]$ADPropsCurrentUser.extensionattribute9
$ReplaceHash['$CurrentUserExtattr10$'] = [string]$ADPropsCurrentUser.extensionattribute10
$ReplaceHash['$CurrentUserExtattr11$'] = [string]$ADPropsCurrentUser.extensionattribute11
$ReplaceHash['$CurrentUserExtattr12$'] = [string]$ADPropsCurrentUser.extensionattribute12
$ReplaceHash['$CurrentUserExtattr13$'] = [string]$ADPropsCurrentUser.extensionattribute13
$ReplaceHash['$CurrentUserExtattr14$'] = [string]$ADPropsCurrentUser.extensionattribute14
$ReplaceHash['$CurrentUserExtattr15$'] = [string]$ADPropsCurrentUser.extensionattribute15
$ReplaceHash['$CurrentUserOffice$'] = [string]$ADPropsCurrentUser.physicaldeliveryofficename
$ReplaceHash['$CurrentUserCompany$'] = [string]$ADPropsCurrentUser.company
$ReplaceHash['$CurrentUserMailnickname$'] = [string]$ADPropsCurrentUser.mailnickname
$ReplaceHash['$CurrentUserDisplayname$'] = [string]$ADPropsCurrentUser.displayname


# Manager of currently logged in user
$ReplaceHash['$CurrentUserManagerGivenname$'] = [string]$ADPropsCurrentUserManager.givenname
$ReplaceHash['$CurrentUserManagerSurname$'] = [string]$ADPropsCurrentUserManager.sn
$ReplaceHash['$CurrentUserManagerDepartment$'] = [string]$ADPropsCurrentUserManager.department
$ReplaceHash['$CurrentUserManagerTitle$'] = [string]$ADPropsCurrentUserManager.title
$ReplaceHash['$CurrentUserManagerStreetaddress$'] = [string]$ADPropsCurrentUserManager.streetaddress
$ReplaceHash['$CurrentUserManagerPostalcode$'] = [string]$ADPropsCurrentUserManager.postalcode
$ReplaceHash['$CurrentUserManagerLocation$'] = [string]$ADPropsCurrentUserManager.l
$ReplaceHash['$CurrentUserManagerCountry$'] = [string]$ADPropsCurrentUserManager.co
$ReplaceHash['$CurrentUserManagerState$'] = [string]$ADPropsCurrentUserManager.st
$ReplaceHash['$CurrentUserManagerTelephone$'] = [string]$ADPropsCurrentUserManager.telephonenumber
$ReplaceHash['$CurrentUserManagerFax$'] = [string]$ADPropsCurrentUserManager.facsimiletelephonenumber
$ReplaceHash['$CurrentUserManagerMobile$'] = [string]$ADPropsCurrentUserManager.mobile
$ReplaceHash['$CurrentUserManagerMail$'] = [string]$ADPropsCurrentUserManager.mail
$ReplaceHash['$CurrentUserManagerPhoto$'] = $ADPropsCurrentUserManager.thumbnailphoto
$ReplaceHash['$CurrentUserManagerPhotodeleteempty$'] = $ADPropsCurrentUserManager.thumbnailphoto
$ReplaceHash['$CurrentUserManagerExtattr1$'] = [string]$ADPropsCurrentUserManager.extensionattribute1
$ReplaceHash['$CurrentUserManagerExtattr2$'] = [string]$ADPropsCurrentUserManager.extensionattribute2
$ReplaceHash['$CurrentUserManagerExtattr3$'] = [string]$ADPropsCurrentUserManager.extensionattribute3
$ReplaceHash['$CurrentUserManagerExtattr4$'] = [string]$ADPropsCurrentUserManager.extensionattribute4
$ReplaceHash['$CurrentUserManagerExtattr5$'] = [string]$ADPropsCurrentUserManager.extensionattribute5
$ReplaceHash['$CurrentUserManagerExtattr6$'] = [string]$ADPropsCurrentUserManager.extensionattribute6
$ReplaceHash['$CurrentUserManagerExtattr7$'] = [string]$ADPropsCurrentUserManager.extensionattribute7
$ReplaceHash['$CurrentUserManagerExtattr8$'] = [string]$ADPropsCurrentUserManager.extensionattribute8
$ReplaceHash['$CurrentUserManagerExtattr9$'] = [string]$ADPropsCurrentUserManager.extensionattribute9
$ReplaceHash['$CurrentUserManagerExtattr10$'] = [string]$ADPropsCurrentUserManager.extensionattribute10
$ReplaceHash['$CurrentUserManagerExtattr11$'] = [string]$ADPropsCurrentUserManager.extensionattribute11
$ReplaceHash['$CurrentUserManagerExtattr12$'] = [string]$ADPropsCurrentUserManager.extensionattribute12
$ReplaceHash['$CurrentUserManagerExtattr13$'] = [string]$ADPropsCurrentUserManager.extensionattribute13
$ReplaceHash['$CurrentUserManagerExtattr14$'] = [string]$ADPropsCurrentUserManager.extensionattribute14
$ReplaceHash['$CurrentUserManagerExtattr15$'] = [string]$ADPropsCurrentUserManager.extensionattribute15
$ReplaceHash['$CurrentUserManagerOffice$'] = [string]$ADPropsCurrentUserManager.physicaldeliveryofficename
$ReplaceHash['$CurrentUserManagerCompany$'] = [string]$ADPropsCurrentUserManager.company
$ReplaceHash['$CurrentUserManagerMailnickname$'] = [string]$ADPropsCurrentUserManager.mailnickname
$ReplaceHash['$CurrentUserManagerDisplayname$'] = [string]$ADPropsCurrentUserManager.displayname


# Current mailbox
$ReplaceHash['$CurrentMailboxGivenname$'] = [string]$ADPropsCurrentMailbox.givenname
$ReplaceHash['$CurrentMailboxSurname$'] = [string]$ADPropsCurrentMailbox.sn
$ReplaceHash['$CurrentMailboxDepartment$'] = [string]$ADPropsCurrentMailbox.department
$ReplaceHash['$CurrentMailboxTitle$'] = [string]$ADPropsCurrentMailbox.title
$ReplaceHash['$CurrentMailboxStreetaddress$'] = [string]$ADPropsCurrentMailbox.streetaddress
$ReplaceHash['$CurrentMailboxPostalcode$'] = [string]$ADPropsCurrentMailbox.postalcode
$ReplaceHash['$CurrentMailboxLocation$'] = [string]$ADPropsCurrentMailbox.l
$ReplaceHash['$CurrentMailboxCountry$'] = [string]$ADPropsCurrentMailbox.co
$ReplaceHash['$CurrentMailboxState$'] = [string]$ADPropsCurrentMailbox.st
$ReplaceHash['$CurrentMailboxTelephone$'] = [string]$ADPropsCurrentMailbox.telephonenumber
$ReplaceHash['$CurrentMailboxFax$'] = [string]$ADPropsCurrentMailbox.facsimiletelephonenumber
$ReplaceHash['$CurrentMailboxMobile$'] = [string]$ADPropsCurrentMailbox.mobile
$ReplaceHash['$CurrentMailboxMail$'] = [string]$ADPropsCurrentMailbox.mail
$ReplaceHash['$CurrentMailboxPhoto$'] = $ADPropsCurrentMailbox.thumbnailphoto
$ReplaceHash['$CurrentMailboxPhotodeleteempty$'] = $ADPropsCurrentMailbox.thumbnailphoto
$ReplaceHash['$CurrentMailboxExtattr1$'] = [string]$ADPropsCurrentMailbox.extensionattribute1
$ReplaceHash['$CurrentMailboxExtattr2$'] = [string]$ADPropsCurrentMailbox.extensionattribute2
$ReplaceHash['$CurrentMailboxExtattr3$'] = [string]$ADPropsCurrentMailbox.extensionattribute3
$ReplaceHash['$CurrentMailboxExtattr4$'] = [string]$ADPropsCurrentMailbox.extensionattribute4
$ReplaceHash['$CurrentMailboxExtattr5$'] = [string]$ADPropsCurrentMailbox.extensionattribute5
$ReplaceHash['$CurrentMailboxExtattr6$'] = [string]$ADPropsCurrentMailbox.extensionattribute6
$ReplaceHash['$CurrentMailboxExtattr7$'] = [string]$ADPropsCurrentMailbox.extensionattribute7
$ReplaceHash['$CurrentMailboxExtattr8$'] = [string]$ADPropsCurrentMailbox.extensionattribute8
$ReplaceHash['$CurrentMailboxExtattr9$'] = [string]$ADPropsCurrentMailbox.extensionattribute9
$ReplaceHash['$CurrentMailboxExtattr10$'] = [string]$ADPropsCurrentMailbox.extensionattribute10
$ReplaceHash['$CurrentMailboxExtattr11$'] = [string]$ADPropsCurrentMailbox.extensionattribute11
$ReplaceHash['$CurrentMailboxExtattr12$'] = [string]$ADPropsCurrentMailbox.extensionattribute12
$ReplaceHash['$CurrentMailboxExtattr13$'] = [string]$ADPropsCurrentMailbox.extensionattribute13
$ReplaceHash['$CurrentMailboxExtattr14$'] = [string]$ADPropsCurrentMailbox.extensionattribute14
$ReplaceHash['$CurrentMailboxExtattr15$'] = [string]$ADPropsCurrentMailbox.extensionattribute15
$ReplaceHash['$CurrentMailboxOffice$'] = [string]$ADPropsCurrentMailbox.physicaldeliveryofficename
$ReplaceHash['$CurrentMailboxCompany$'] = [string]$ADPropsCurrentMailbox.company
$ReplaceHash['$CurrentMailboxMailnickname$'] = [string]$ADPropsCurrentMailbox.mailnickname
$ReplaceHash['$CurrentMailboxDisplayname$'] = [string]$ADPropsCurrentMailbox.displayname


# Manager of current mailbox
$ReplaceHash['$CurrentMailboxManagerGivenname$'] = [string]$ADPropsCurrentMailboxManager.givenname
$ReplaceHash['$CurrentMailboxManagerSurname$'] = [string]$ADPropsCurrentMailboxManager.sn
$ReplaceHash['$CurrentMailboxManagerDepartment$'] = [string]$ADPropsCurrentMailboxManager.department
$ReplaceHash['$CurrentMailboxManagerTitle$'] = [string]$ADPropsCurrentMailboxManager.title
$ReplaceHash['$CurrentMailboxManagerStreetaddress$'] = [string]$ADPropsCurrentMailboxManager.streetaddress
$ReplaceHash['$CurrentMailboxManagerPostalcode$'] = [string]$ADPropsCurrentMailboxManager.postalcode
$ReplaceHash['$CurrentMailboxManagerLocation$'] = [string]$ADPropsCurrentMailboxManager.l
$ReplaceHash['$CurrentMailboxManagerCountry$'] = [string]$ADPropsCurrentMailboxManager.co
$ReplaceHash['$CurrentMailboxManagerState$'] = [string]$ADPropsCurrentMailboxManager.st
$ReplaceHash['$CurrentMailboxManagerTelephone$'] = [string]$ADPropsCurrentMailboxManager.telephonenumber
$ReplaceHash['$CurrentMailboxManagerFax$'] = [string]$ADPropsCurrentMailboxManager.facsimiletelephonenumber
$ReplaceHash['$CurrentMailboxManagerMobile$'] = [string]$ADPropsCurrentMailboxManager.mobile
$ReplaceHash['$CurrentMailboxManagerMail$'] = [string]$ADPropsCurrentMailboxManager.mail
$ReplaceHash['$CurrentMailboxManagerPhoto$'] = $ADPropsCurrentMailboxManager.thumbnailphoto
$ReplaceHash['$CurrentMailboxManagerPhotodeleteempty$'] = $ADPropsCurrentMailboxManager.thumbnailphoto
$ReplaceHash['$CurrentMailboxManagerExtattr1$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute1
$ReplaceHash['$CurrentMailboxManagerExtattr2$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute2
$ReplaceHash['$CurrentMailboxManagerExtattr3$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute3
$ReplaceHash['$CurrentMailboxManagerExtattr4$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute4
$ReplaceHash['$CurrentMailboxManagerExtattr5$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute5
$ReplaceHash['$CurrentMailboxManagerExtattr6$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute6
$ReplaceHash['$CurrentMailboxManagerExtattr7$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute7
$ReplaceHash['$CurrentMailboxManagerExtattr8$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute8
$ReplaceHash['$CurrentMailboxManagerExtattr9$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute9
$ReplaceHash['$CurrentMailboxManagerExtattr10$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute10
$ReplaceHash['$CurrentMailboxManagerExtattr11$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute11
$ReplaceHash['$CurrentMailboxManagerExtattr12$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute12
$ReplaceHash['$CurrentMailboxManagerExtattr13$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute13
$ReplaceHash['$CurrentMailboxManagerExtattr14$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute14
$ReplaceHash['$CurrentMailboxManagerExtattr15$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute15
$ReplaceHash['$CurrentMailboxManagerOffice$'] = [string]$ADPropsCurrentMailboxManager.physicaldeliveryofficename
$ReplaceHash['$CurrentMailboxManagerCompany$'] = [string]$ADPropsCurrentMailboxManager.company
$ReplaceHash['$CurrentMailboxManagerMailnickname$'] = [string]$ADPropsCurrentMailboxManager.mailnickname
$ReplaceHash['$CurrentMailboxManagerDisplayname$'] = [string]$ADPropsCurrentMailboxManager.displayname


# $CurrentUserNamewithtitles$, $CurrentUserManagerNamewithtitles$
# $CurrentMailboxNamewithtitles$, $CurrentMailboxManagerNamewithtitles$
# Academic titles according to standards in German speaking countries
# <custom AD attribute 'svstitelvorne'> <standard AD attribute 'givenname'> <standard AD attribute 'surname'>, <custom AD attribute 'svstitelhinten'>
# If one or more attributes are not set, unnecessary whitespaces and commas are avoided by using '-join'
# Examples:
#   Mag. Dr. John Doe, BA MA PhD
#   Dr. John Doe
#   John Doe, PhD
#   John Doe
$ReplaceHash['$CurrentUserNamewithtitles$'] = (((((([string]$ADPropsCurrentUser.svstitelvorne, [string]$ADPropsCurrentUser.givenname, [string]$ADPropsCurrentUser.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUser.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CurrentUserManagerNamewithtitles$'] = (((((([string]$ADPropsCurrentUserManager.svstitelvorne, [string]$ADPropsCurrentUserManager.givenname, [string]$ADPropsCurrentUserManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUserManager.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CurrentMailboxNamewithtitles$'] = (((((([string]$ADPropsCurrentMailbox.svstitelvorne, [string]$ADPropsCurrentMailbox.givenname, [string]$ADPropsCurrentMailbox.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailbox.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CurrentMailboxManagerNamewithtitles$'] = (((((([string]$ADPropsCurrentMailboxManager.svstitelvorne, [string]$ADPropsCurrentMailboxManager.givenname, [string]$ADPropsCurrentMailboxManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailboxManager.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')
