# This file allows defining custom replacement variables for Set-OutlookSignatures
#
# This script is executed as a whole once for each mailbox.
# It allows for complex replacement variable handling (complex string transformations, retrieving information from web services and databases, etc.).
#
# Attention: The configuration file is executed as part of Set-OutlookSignatures.ps1 and is not checked for any harmful content. Please only allow qualified technicians write access to this file, only use it to to define replacement variables and test it thoroughly.
#
# Replacement variable names are case sensitive.
# It is required to use full uppercase replacement variable names.
#
# Active Directory property names are case sensitive.
# It is required to use full lowercase Active Directory property names.
#
# A variable defined in this file overrides the definition of the same variable defined earlier in the script.
#
#
# What is the recommended approach for custom configuration files?
# You should not change the default configuration file `'.\config\default replacement variable.ps1'`, as it might be changed in a future release of Set-OutlookSignatures. In this case, you would have to sort out the changes yourself.
#
# The following steps are recommended:
# 1. Create a new custom configuration file in a separate folder.
# 2. The first step in the new custom configuration file should be to load the default configuration file:
#    # Loading default replacement variables shipped with Set-OutlookSignatures
#    . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath '\\server\share\folder\Set-OutlookSignatures\config\default replacement variables.ps1' -Raw)))
# 3. After importing the default configuration file, existing replacement variables can be altered with custom definitions and new replacement variables can be added.
# 4. Instead of altering existing replacement variables, it is recommended to create new replacement variables with modified content.
# 5. Start Set-OutlookSignatures with the parameter 'ReplacementVariableConfigFile' pointing to the new custom configuration file.


# Currently logged on user
$ReplaceHash['$CURRENTUSERGIVENNAME$'] = [string]$ADPropsCurrentUser.givenname
$ReplaceHash['$CURRENTUSERSURNAME$'] = [string]$ADPropsCurrentUser.sn
$ReplaceHash['$CURRENTUSERDEPARTMENT$'] = [string]$ADPropsCurrentUser.department
$ReplaceHash['$CURRENTUSERTITLE$'] = [string]$ADPropsCurrentUser.title
$ReplaceHash['$CURRENTUSERSTREETADDRESS$'] = [string]$ADPropsCurrentUser.streetaddress
$ReplaceHash['$CURRENTUSERPOSTALCODE$'] = [string]$ADPropsCurrentUser.postalcode
$ReplaceHash['$CURRENTUSERLOCATION$'] = [string]$ADPropsCurrentUser.l
$ReplaceHash['$CURRENTUSERCOUNTRY$'] = [string]$ADPropsCurrentUser.co
$ReplaceHash['$CURRENTUSERTELEPHONE$'] = [string]$ADPropsCurrentUser.telephonenumber
$ReplaceHash['$CURRENTUSERFAX$'] = [string]$ADPropsCurrentUser.facsimiletelephonenumber
$ReplaceHash['$CURRENTUSERMOBILE$'] = [string]$ADPropsCurrentUser.mobile
$ReplaceHash['$CURRENTUSERMAIL$'] = [string]$ADPropsCurrentUser.mail
$ReplaceHash['$CURRENTUSERPHOTO$'] = $ADPropsCurrentUser.thumbnailphoto
$ReplaceHash['$CURRENTUSERPHOTODELETEEMPTY$'] = $ADPropsCurrentUser.thumbnailphoto
$ReplaceHash['$CURRENTUSEREXTATTR1$'] = [string]$ADPropsCurrentUser.extensionattribute1
$ReplaceHash['$CURRENTUSEREXTATTR2$'] = [string]$ADPropsCurrentUser.extensionattribute2
$ReplaceHash['$CURRENTUSEREXTATTR3$'] = [string]$ADPropsCurrentUser.extensionattribute3
$ReplaceHash['$CURRENTUSEREXTATTR4$'] = [string]$ADPropsCurrentUser.extensionattribute4
$ReplaceHash['$CURRENTUSEREXTATTR5$'] = [string]$ADPropsCurrentUser.extensionattribute5
$ReplaceHash['$CURRENTUSEREXTATTR6$'] = [string]$ADPropsCurrentUser.extensionattribute6
$ReplaceHash['$CURRENTUSEREXTATTR7$'] = [string]$ADPropsCurrentUser.extensionattribute7
$ReplaceHash['$CURRENTUSEREXTATTR8$'] = [string]$ADPropsCurrentUser.extensionattribute8
$ReplaceHash['$CURRENTUSEREXTATTR9$'] = [string]$ADPropsCurrentUser.extensionattribute9
$ReplaceHash['$CURRENTUSEREXTATTR10$'] = [string]$ADPropsCurrentUser.extensionattribute10
$ReplaceHash['$CURRENTUSEREXTATTR11$'] = [string]$ADPropsCurrentUser.extensionattribute11
$ReplaceHash['$CURRENTUSEREXTATTR12$'] = [string]$ADPropsCurrentUser.extensionattribute12
$ReplaceHash['$CURRENTUSEREXTATTR13$'] = [string]$ADPropsCurrentUser.extensionattribute13
$ReplaceHash['$CURRENTUSEREXTATTR14$'] = [string]$ADPropsCurrentUser.extensionattribute14
$ReplaceHash['$CURRENTUSEREXTATTR15$'] = [string]$ADPropsCurrentUser.extensionattribute15
$ReplaceHash['$CURRENTUSEROFFICE$'] = [string]$ADPropsCurrentUser.physicaldeliveryofficename
$ReplaceHash['$CURRENTUSERCOMPANY$'] = [string]$ADPropsCurrentUser.company
$ReplaceHash['$CURRENTUSERMAILNICKNAME$'] = [string]$ADPropsCurrentUser.mailnickname
$ReplaceHash['$CURRENTUSERDISPLAYNAME$'] = [string]$ADPropsCurrentUser.displayname


# Manager of currently logged on user
$ReplaceHash['$CURRENTUSERMANAGERGIVENNAME$'] = [string]$ADPropsCurrentUserManager.givenname
$ReplaceHash['$CURRENTUSERMANAGERSURNAME$'] = [string]$ADPropsCurrentUserManager.sn
$ReplaceHash['$CURRENTUSERMANAGERDEPARTMENT$'] = [string]$ADPropsCurrentUserManager.department
$ReplaceHash['$CURRENTUSERMANAGERTITLE$'] = [string]$ADPropsCurrentUserManager.title
$ReplaceHash['$CURRENTUSERMANAGERSTREETADDRESS$'] = [string]$ADPropsCurrentUserManager.streetaddress
$ReplaceHash['$CURRENTUSERMANAGERPOSTALCODE$'] = [string]$ADPropsCurrentUserManager.postalcode
$ReplaceHash['$CURRENTUSERMANAGERLOCATION$'] = [string]$ADPropsCurrentUserManager.l
$ReplaceHash['$CURRENTUSERMANAGERCOUNTRY$'] = [string]$ADPropsCurrentUserManager.co
$ReplaceHash['$CURRENTUSERMANAGERTELEPHONE$'] = [string]$ADPropsCurrentUserManager.telephonenumber
$ReplaceHash['$CURRENTUSERMANAGERFAX$'] = [string]$ADPropsCurrentUserManager.facsimiletelephonenumber
$ReplaceHash['$CURRENTUSERMANAGERMOBILE$'] = [string]$ADPropsCurrentUserManager.mobile
$ReplaceHash['$CURRENTUSERMANAGERMAIL$'] = [string]$ADPropsCurrentUserManager.mail
$ReplaceHash['$CURRENTUSERMANAGERPHOTO$'] = $ADPropsCurrentUserManager.thumbnailphoto
$ReplaceHash['$CURRENTUSERMANAGERPHOTODELETEEMPTY$'] = $ADPropsCurrentUserManager.thumbnailphoto
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR1$'] = [string]$ADPropsCurrentUserManager.extensionattribute1
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR2$'] = [string]$ADPropsCurrentUserManager.extensionattribute2
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR3$'] = [string]$ADPropsCurrentUserManager.extensionattribute3
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR4$'] = [string]$ADPropsCurrentUserManager.extensionattribute4
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR5$'] = [string]$ADPropsCurrentUserManager.extensionattribute5
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR6$'] = [string]$ADPropsCurrentUserManager.extensionattribute6
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR7$'] = [string]$ADPropsCurrentUserManager.extensionattribute7
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR8$'] = [string]$ADPropsCurrentUserManager.extensionattribute8
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR9$'] = [string]$ADPropsCurrentUserManager.extensionattribute9
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR10$'] = [string]$ADPropsCurrentUserManager.extensionattribute10
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR11$'] = [string]$ADPropsCurrentUserManager.extensionattribute11
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR12$'] = [string]$ADPropsCurrentUserManager.extensionattribute12
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR13$'] = [string]$ADPropsCurrentUserManager.extensionattribute13
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR14$'] = [string]$ADPropsCurrentUserManager.extensionattribute14
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR15$'] = [string]$ADPropsCurrentUserManager.extensionattribute15
$ReplaceHash['$CURRENTUSERMANAGEROFFICE$'] = [string]$ADPropsCurrentUserManager.physicaldeliveryofficename
$ReplaceHash['$CURRENTUSERMANAGERCOMPANY$'] = [string]$ADPropsCurrentUserManager.company
$ReplaceHash['$CURRENTUSERMANAGERMAILNICKNAME$'] = [string]$ADPropsCurrentUserManager.mailnickname
$ReplaceHash['$CURRENTUSERMANAGERDISPLAYNAME$'] = [string]$ADPropsCurrentUserManager.displayname


# Current mailbox
$ReplaceHash['$CURRENTMAILBOXGIVENNAME$'] = [string]$ADPropsCurrentMailbox.givenname
$ReplaceHash['$CURRENTMAILBOXSURNAME$'] = [string]$ADPropsCurrentMailbox.sn
$ReplaceHash['$CURRENTMAILBOXDEPARTMENT$'] = [string]$ADPropsCurrentMailbox.department
$ReplaceHash['$CURRENTMAILBOXTITLE$'] = [string]$ADPropsCurrentMailbox.title
$ReplaceHash['$CURRENTMAILBOXSTREETADDRESS$'] = [string]$ADPropsCurrentMailbox.streetaddress
$ReplaceHash['$CURRENTMAILBOXPOSTALCODE$'] = [string]$ADPropsCurrentMailbox.postalcode
$ReplaceHash['$CURRENTMAILBOXLOCATION$'] = [string]$ADPropsCurrentMailbox.l
$ReplaceHash['$CURRENTMAILBOXCOUNTRY$'] = [string]$ADPropsCurrentMailbox.co
$ReplaceHash['$CURRENTMAILBOXTELEPHONE$'] = [string]$ADPropsCurrentMailbox.telephonenumber
$ReplaceHash['$CURRENTMAILBOXFAX$'] = [string]$ADPropsCurrentMailbox.facsimiletelephonenumber
$ReplaceHash['$CURRENTMAILBOXMOBILE$'] = [string]$ADPropsCurrentMailbox.mobile
$ReplaceHash['$CURRENTMAILBOXMAIL$'] = [string]$ADPropsCurrentMailbox.mail
$ReplaceHash['$CURRENTMAILBOXPHOTO$'] = $ADPropsCurrentMailbox.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXPHOTODELETEEMPTY$'] = $ADPropsCurrentMailbox.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXEXTATTR1$'] = [string]$ADPropsCurrentMailbox.extensionattribute1
$ReplaceHash['$CURRENTMAILBOXEXTATTR2$'] = [string]$ADPropsCurrentMailbox.extensionattribute2
$ReplaceHash['$CURRENTMAILBOXEXTATTR3$'] = [string]$ADPropsCurrentMailbox.extensionattribute3
$ReplaceHash['$CURRENTMAILBOXEXTATTR4$'] = [string]$ADPropsCurrentMailbox.extensionattribute4
$ReplaceHash['$CURRENTMAILBOXEXTATTR5$'] = [string]$ADPropsCurrentMailbox.extensionattribute5
$ReplaceHash['$CURRENTMAILBOXEXTATTR6$'] = [string]$ADPropsCurrentMailbox.extensionattribute6
$ReplaceHash['$CURRENTMAILBOXEXTATTR7$'] = [string]$ADPropsCurrentMailbox.extensionattribute7
$ReplaceHash['$CURRENTMAILBOXEXTATTR8$'] = [string]$ADPropsCurrentMailbox.extensionattribute8
$ReplaceHash['$CURRENTMAILBOXEXTATTR9$'] = [string]$ADPropsCurrentMailbox.extensionattribute9
$ReplaceHash['$CURRENTMAILBOXEXTATTR10$'] = [string]$ADPropsCurrentMailbox.extensionattribute10
$ReplaceHash['$CURRENTMAILBOXEXTATTR11$'] = [string]$ADPropsCurrentMailbox.extensionattribute11
$ReplaceHash['$CURRENTMAILBOXEXTATTR12$'] = [string]$ADPropsCurrentMailbox.extensionattribute12
$ReplaceHash['$CURRENTMAILBOXEXTATTR13$'] = [string]$ADPropsCurrentMailbox.extensionattribute13
$ReplaceHash['$CURRENTMAILBOXEXTATTR14$'] = [string]$ADPropsCurrentMailbox.extensionattribute14
$ReplaceHash['$CURRENTMAILBOXEXTATTR15$'] = [string]$ADPropsCurrentMailbox.extensionattribute15
$ReplaceHash['$CURRENTMAILBOXOFFICE$'] = [string]$ADPropsCurrentMailbox.physicaldeliveryofficename
$ReplaceHash['$CURRENTMAILBOXCOMPANY$'] = [string]$ADPropsCurrentMailbox.company
$ReplaceHash['$CURRENTMAILBOXMAILNICKNAME$'] = [string]$ADPropsCurrentMailbox.mailnickname
$ReplaceHash['$CURRENTMAILBOXDISPLAYNAME$'] = [string]$ADPropsCurrentMailbox.displayname


# Manager of current mailbox
$ReplaceHash['$CURRENTMAILBOXMANAGERGIVENNAME$'] = [string]$ADPropsCurrentMailboxManager.givenname
$ReplaceHash['$CURRENTMAILBOXMANAGERSURNAME$'] = [string]$ADPropsCurrentMailboxManager.sn
$ReplaceHash['$CURRENTMAILBOXMANAGERDEPARTMENT$'] = [string]$ADPropsCurrentMailboxManager.department
$ReplaceHash['$CURRENTMAILBOXMANAGERTITLE$'] = [string]$ADPropsCurrentMailboxManager.title
$ReplaceHash['$CURRENTMAILBOXMANAGERSTREETADDRESS$'] = [string]$ADPropsCurrentMailboxManager.streetaddress
$ReplaceHash['$CURRENTMAILBOXMANAGERPOSTALCODE$'] = [string]$ADPropsCurrentMailboxManager.postalcode
$ReplaceHash['$CURRENTMAILBOXMANAGERLOCATION$'] = [string]$ADPropsCurrentMailboxManager.l
$ReplaceHash['$CURRENTMAILBOXMANAGERCOUNTRY$'] = [string]$ADPropsCurrentMailboxManager.co
$ReplaceHash['$CURRENTMAILBOXMANAGERTELEPHONE$'] = [string]$ADPropsCurrentMailboxManager.telephonenumber
$ReplaceHash['$CURRENTMAILBOXMANAGERFAX$'] = [string]$ADPropsCurrentMailboxManager.facsimiletelephonenumber
$ReplaceHash['$CURRENTMAILBOXMANAGERMOBILE$'] = [string]$ADPropsCurrentMailboxManager.mobile
$ReplaceHash['$CURRENTMAILBOXMANAGERMAIL$'] = [string]$ADPropsCurrentMailboxManager.mail
$ReplaceHash['$CURRENTMAILBOXMANAGERPHOTO$'] = $ADPropsCurrentMailboxManager.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$'] = $ADPropsCurrentMailboxManager.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR1$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute1
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR2$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute2
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR3$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute3
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR4$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute4
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR5$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute5
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR6$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute6
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR7$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute7
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR8$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute8
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR9$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute9
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR10$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute10
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR11$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute11
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR12$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute12
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR13$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute13
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR14$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute14
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR15$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute15
$ReplaceHash['$CURRENTMAILBOXMANAGEROFFICE$'] = [string]$ADPropsCurrentMailboxManager.physicaldeliveryofficename
$ReplaceHash['$CURRENTMAILBOXMANAGERCOMPANY$'] = [string]$ADPropsCurrentMailboxManager.company
$ReplaceHash['$CURRENTMAILBOXMANAGERMAILNICKNAME$'] = [string]$ADPropsCurrentMailboxManager.mailnickname
$ReplaceHash['$CURRENTMAILBOXMANAGERDISPLAYNAME$'] = [string]$ADPropsCurrentMailboxManager.displayname


# $CURRENTUSERNAMEWITHTITLES$, $CURRENTUSERMANAGERNAMEWITHTITLES$
# $CURRENTMAILBOXNAMEWITHTITLES$, $CURRENTMAILBOXMANAGERNAMEWITHTITLES$
# Academic titles according to standards in German speaking countries
# <custom AD attribute 'svstitelvorne'> <standard AD attribute 'givenname'> <standard AD attribute 'surname'>, <custom AD attribute 'svstitelhinten'>
# If one or more attributes are not set, unnecessary whitespaces and commas are avoided by using '-join'
# Examples:
#   Mag. Dr. John Doe, BA MA PhD
#   Dr. John Doe
#   John Doe, PhD
#   John Doe
$ReplaceHash['$CURRENTUSERNAMEWITHTITLES$'] = (((((([string]$ADPropsCurrentUser.svstitelvorne, [string]$ADPropsCurrentUser.givenname, [string]$ADPropsCurrentUser.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUser.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CURRENTUSERMANAGERNAMEWITHTITLES$'] = (((((([string]$ADPropsCurrentUserManager.svstitelvorne, [string]$ADPropsCurrentUserManager.givenname, [string]$ADPropsCurrentUserManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUserManager.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CURRENTMAILBOXNAMEWITHTITLES$'] = (((((([string]$ADPropsCurrentMailbox.svstitelvorne, [string]$ADPropsCurrentMailbox.givenname, [string]$ADPropsCurrentMailbox.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailbox.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CURRENTMAILBOXMANAGERNAMEWITHTITLES$'] = (((((([string]$ADPropsCurrentMailboxManager.svstitelvorne, [string]$ADPropsCurrentMailboxManager.givenname, [string]$ADPropsCurrentMailboxManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailboxManager.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')
