<!-- omit in toc -->
# <a href="https://github.com/GruberMarkus/Set-OutlookSignatures"><img src="../src/logo/Set-OutlookSignatures%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-outlookSignatures"></a><br>Centrally&nbsp;manage&nbsp;and&nbsp;deploy Outlook&nbsp;text&nbsp;signatures&nbsp;and Out&nbsp;of&nbsp;Office&nbsp;auto&nbsp;reply&nbsp;messages.<br><a href="https://github.com/GruberMarkus/Set-OutlookSignatures/blob/main/license.txt"><img src="https://img.shields.io/github/license/grubermarkus/Set-OutlookSignatures" alt=""></a> <a href="https://www.paypal.com/donate?business=JBM584K3L5PX4&no_recurring=0&currency_code=EUR"><img src="https://img.shields.io/badge/sponsor-grey?logo=paypal" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/clones.svg" alt="" data-external="1">

<object data="/github/forks/badges/shields?label=Fork&amp;style=social"></object>


<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases"><img src="https://img.shields.io/github/downloads/GruberMarkus/Set-OutlookSignatures/total" alt="" data-external="1"></a> <a href="https://github.com/grubermarkus/set-outlooksignatures/stargazers"><img src="https://img.shields.io/github/stars/grubermarkus/set-outlooksignatures" alt="" data-external="1"></a> <a href="https://github.com/grubermarkus/set-outlooksignatures/issues"><img src="https://img.shields.io/github/issues/grubermarkus/set-outlooksignatures" alt="" data-external="1"></a>  

# Features <!-- omit in toc -->
**Signatures and OOF messages can be:**
- Generated from templates in DOCX or HTML file format  
- Customized with a broad range of variables, including photos, from Active Directory and other sources  
- Applied to all mailboxes (including shared mailboxes), specific mailbox groups or specific email addresses, for every primary mailbox across all Outlook profiles  
- Assigned time ranges within which they are valid  
- Set as default signature for new mails, or for replies and forwards (signatures only)  
- Set as default OOF message for internal or external recipients (OOF messages only)  
- Set in Outlook Web for the currently logged-on user  
- Centrally managed only or exist along user created signatures (signatures only)  
- Copied to an alternate path for easy access on mobile devices not directly supported by this script (signatures only)
  
**Sample templates** for signatures and OOF messages demonstrate all available features and are provided as .docx and .htm files.

**Simulation mode** allows content creators and admins to simulate the behavior of the script and to inspect the resulting signature files before going live.
  
The script is **designed to work in big and complex environments** (Exchange resource forest scenarios, across AD trusts, multi-level AD subdomains, many objects). The script is **multi-client capable** by using different template paths, configuration files and script parameters.
  
The script is **Free and Open-Source Software (FOSS)**. It is published under the MIT license which is approved, among others, by the Free Software Foundation (FSF), the Open Source Initiative (OSI) and is compatible with the General Public License (GPL) v3. Please see `'.\docs\LICENSE.txt'` for copyright and MIT license details.
# Table of Contents <!-- omit in toc -->
- [1. Requirements](#1-requirements)
- [2. Parameters](#2-parameters)
  - [2.1. SignatureTemplatePath](#21-signaturetemplatepath)
  - [2.2. ReplacementVariableConfigFile](#22-replacementvariableconfigfile)
  - [2.3. DomainsToCheckForGroups](#23-domainstocheckforgroups)
  - [2.4. DeleteUserCreatedSignatures](#24-deleteusercreatedsignatures)
  - [2.5. SetCurrentUserOutlookWebSignature](#25-setcurrentuseroutlookwebsignature)
  - [2.6. SetCurrentUserOOFMessage](#26-setcurrentuseroofmessage)
  - [2.7. OOFTemplatePath](#27-ooftemplatepath)
  - [2.8. AdditionalSignaturePath](#28-additionalsignaturepath)
  - [2.9. AdditionalSignaturePathFolder](#29-additionalsignaturepathfolder)
  - [2.10. UseHtmTemplates](#210-usehtmtemplates)
  - [2.11. SimulationUser](#211-simulationuser)
  - [2.12. SimulationMailboxes](#212-simulationmailboxes)
- [3. Outlook signature path](#3-outlook-signature-path)
- [4. Mailboxes](#4-mailboxes)
- [5. Group membership](#5-group-membership)
- [6. Removing old signatures](#6-removing-old-signatures)
- [7. Error handling](#7-error-handling)
- [8. Run script while Outlook is running](#8-run-script-while-outlook-is-running)
- [9. Signature and OOF file format](#9-signature-and-oof-file-format)
  - [9.1. Signature and OOF file naming](#91-signature-and-oof-file-naming)
  - [9.2. Allowed filename tags](#92-allowed-filename-tags)
- [10. Signature and OOF application order](#10-signature-and-oof-application-order)
- [11. Variable replacement](#11-variable-replacement)
  - [11.1. Photos from Active Directory](#111-photos-from-active-directory)
- [12. Outlook Web](#12-outlook-web)
- [13. Simulation mode](#13-simulation-mode)
- [14. FAQ](#14-faq)
  - [14.1. Where can I find the changelog?](#141-where-can-i-find-the-changelog)
  - [14.2. How can I contribute, propose a new feature or file a bug?](#142-how-can-i-contribute-propose-a-new-feature-or-file-a-bug)
  - [14.3. Why use legacyExchangeDN to find the user behind a mailbox, and not mail or proxyAddresses?](#143-why-use-legacyexchangedn-to-find-the-user-behind-a-mailbox-and-not-mail-or-proxyaddresses)
  - [14.4. How is the personal mailbox of the currently logged-on user identified?](#144-how-is-the-personal-mailbox-of-the-currently-logged-on-user-identified)
  - [14.5. Which ports are required?](#145-which-ports-are-required)
  - [14.6. Why is Out of Office abbreviated OOF and not OOO?](#146-why-is-out-of-office-abbreviated-oof-and-not-ooo)
  - [14.7. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.](#147-should-i-use-docx-or-htm-as-file-format-for-templates-signatures-in-outlook-sometimes-look-different-than-my-templates)
  - [14.8. How can I log the script output?](#148-how-can-i-log-the-script-output)
  - [14.9. Can multiple script instances run in parallel?](#149-can-multiple-script-instances-run-in-parallel)
  - [14.10. How do I start the script from the command line or a scheduled task?](#1410-how-do-i-start-the-script-from-the-command-line-or-a-scheduled-task)
  - [14.11. How to create a shortcut to the script with parameters?](#1411-how-to-create-a-shortcut-to-the-script-with-parameters)
  - [14.12. What is the recommended approach for implementing the software?](#1412-what-is-the-recommended-approach-for-implementing-the-software)
  - [14.13. What about the new signature roaming feature Microsoft announced?](#1413-what-about-the-new-signature-roaming-feature-microsoft-announced)
  
# 1. Requirements  
Requires Outlook and Word, at least version 2010.  
The script must run in the security context of the currently logged-on user.

The script must run in PowerShell Full Language mode. Constrained Language mode is not supported, as some features such as BASE64 conversions are not available in this mode or require very slow workarounds. If you use AppLocker or a comparable solution, you may need to digitally sign the PowerShell script.

The paths to the template files (SignatureTemplatePath, OOFTemplatePath) must be accessible by the currently logged-on user. The template files must be at least readable for the currently logged-on user.  
# 2. Parameters  
## 2.1. SignatureTemplatePath  
The parameter SignatureTemplatePath tells the script where signature template files are stored.

Local and remote paths are supported. Local paths can be absolute (`'C:\Signature templates'`) or relative to the script path (`'.\templates\Signatures'`).

WebDAV paths are supported (https only): `'https://server.domain/SignatureSite/SignatureTemplates'` or `'\\server.domain@SSL\SignatureSite\SignatureTemplates'`

The currently logged-on user needs at least read access to the path.

Default value: `'.\templates\Signatures DOCX'`  
## 2.2. ReplacementVariableConfigFile  
The parameter ReplacementVariableConfigFile tells the script where the file defining replacement variables is located.

Local and remote paths are supported. Local paths can be absolute (`'C:\config\default replacement variables.ps1'`) or relative to the script path (`'.\config\default replacement variables.ps1'`).

WebDAV paths are supported (https only): `'https://server.domain/SignatureSite/config/default replacement variables.ps1'` or `'\\server.domain@SSL\SignatureSite\config\default replacement variables.ps1'`

The currently logged-on user needs at least read access to the file.

Default value: `'.\config\default replacement variables.ps1'`  
## 2.3. DomainsToCheckForGroups  
The parameters tells the script which domains should be used to search for mailbox and user group membership.

The default value, `'\*'` tells the script to query all trusted domains in the Active Directory forest of the logged-on user.

For a custom list of domains/forests, specify them as comma-separated list of strings: `"domain-a.local", "dc=example,dc=com", "domain-b.internal"`.

When a domain/forest in the custom list starts with a dash or minus (`'-domain-a.local'`), this domain is removed from the list.

The `'\*'` entry in a custom list is only considered when it is the first entry of the list.

The Active Directory forest of the currently logged-on user is always considered.

Default value: `'*'`  
## 2.4. DeleteUserCreatedSignatures  
Shall the script delete signatures which were created by the user itself? The default value for this parameter is `$false`.

Remark: The script always deletes signatures which were deployed by the script earlier, but are no longer available in the central repository.

Default value: `$false`  
## 2.5. SetCurrentUserOutlookWebSignature  
Shall the script set the Outlook Web signature of the currently logged on user?

Default value: `$true`  
## 2.6. SetCurrentUserOOFMessage  
Shall the script set the Out of Office (OOF) auto reply message of the currently logged on user?

Default value: `$true`  
## 2.7. OOFTemplatePath  
Path to centrally managed Out of Office (OOF) auto reply templates.

Local and remote paths are supported.

Local paths can be absolute (`'C:\OOF templates'`) or relative to the script path (`'.\templates\Out of Office'`).

WebDAV paths are supported (https only): `'https://server.domain/SignatureSite/OOFTemplates'` or `'\\server.domain@SSL\SignatureSite\OOFTemplates'`

The currently logged-on user needs at least read access to the path.

Default value: `'.\templates\Out of Office DOCX'`  
## 2.8. AdditionalSignaturePath  
An additional path that the signatures shall be copied to.  
Ideally, this path is available on all devices of the user, for example via Microsoft OneDrive or Nextcloud.

This way, the user can easily copy-paste his preferred preconfigured signature for use in a mail app not supported by this script, such as Microsoft Outlook Mobile, Apple Mail, Google Gmail or Samsung Email.

Local and remote paths are supported.

Local paths can be absolute (`'C:\Outlook signatures'`) or relative to the script path (`'.\Outlook signatures'`).

WebDAV paths are supported (https only): `'https://server.domain/User/Outlook signatures'` or `'\\server.domain@SSL\User\Outlook signatures'`

The currently logged-on user needs at least write access to the path.

Default value: `"$([environment]::GetFolderPath("MyDocuments"))\Outlook signatures"`  
## 2.9. AdditionalSignaturePathFolder
A folder or folder structure below AdditionalSignaturePath.  
If the folder or folder structure does not exist, it is created.

Default value: `'Outlook signatures'`  
## 2.10. UseHtmTemplates  
With this parameter, the script searches for templates with the extension .htm instead of .docx.

Each format has advantages and disadvantages, please see "[13.5. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.](#135-should-i-use-docx-or-htm-as-file-format-for-templates-signatures-in-outlook-sometimes-look-different-than-my-templates)" for a quick overview.

Default value: `$false`  
## 2.11. SimulationUser  
SimulationUser is a mandatory parameter for simulation mode. This value replaces the currently logged-on user.

See "[13. Simulation mode](#13-simulation-mode)" for details.  
## 2.12. SimulationMailboxes  
SimulationMailboxes is optional for simulation mode, although highly recommended. It is a comma separated list of strings replacing the list of mailboxes otherwise gathered from the registry.

See "[13. Simulation mode](#13-simulation-mode)" for details.  
# 3. Outlook signature path  
The Outlook signature path is retrieved from the users registry, so the script is language independent.

The registry setting does not allow for absolute paths, only for paths relative to `'%APPDATA%\Microsoft'`.

If the relative path set in the registry would be a valid path but does not exist, the script creates it.  
# 4. Mailboxes  
The script only considers primary mailboxes, these are mailboxes added as separate accounts.

This is the same way Outlook handles mailboxes from a signature perspective: Outlook can not handle signatures for non-primary mailboxes (added via "Open these additional mailboxes").

The script is created for Exchange environments. Non-Exchange mailboxes can not have OOF messages or group signatures, but common and mailbox specific signatures.  
# 5. Group membership  
The script considers all groups the currently logged-on user belongs to, as well as all groups the currently processed mailbox belongs to.

For both sets of groups, group membership is evaluated against the whole Active Directory forest of the currently logged-on user, and against all trusted domains the user has access to.

The script works fine with linked mailboxes in Exchange resource forest scenarios.

Trusted domains can be modified with the DomainsToCheckForGroups parameter.

Group membership is achieved by querying the tokenGroups attribute, which is not only very fast and resource saving on client and server, but also considers sIDHistory.  
# 6. Removing old signatures  
The script always deletes signatures which were deployed by the script earlier, but are no longer available in the central repository. The script marks each processed signature with a specific HTML tag, which enables this cleaning feature.

Signatures created manually by the user are not deleted by default, this behavior can be changed with the DeleteUserCreatedSignatures parameter.  
# 7. Error handling  
Error handling is implemented rudimentarily.  
# 8. Run script while Outlook is running  
Outlook and the script can run simultaneously.

New and changed signatures can be used instantly in Outlook.

Changing which signature is to be used as default signature for new mails or for replies and forwards requires restarting Outlook.   
# 9. Signature and OOF file format  
Only Word files with the extension .docx and HTML files with the extension .htm are supported as signature and OOF template files.  
## 9.1. Signature and OOF file naming  
The script copies every signature and OOF file as-is, with one exception: When tags are defined in the file name, these tags are removed.

Tags must be placed before the file extension and be separated from the base filename with a period.

Examples:  
- `'Company external German.docx'` -> `'Company external German.htm'`, no changes  
- `'Company external German.[defaultNew].docx'` -> `'Company external German.htm'`, tag(s) is/are removed  
- `'Company external [English].docx'` -> `'Company external [English].htm'`, tag(s) is/are not removed, because there is no dot before  
- `'Company external [English].[defaultNew] [Company-AD All Employees].docx'` -> `'Company external [English].htm'`, tag(s) is/are removed, because they are separated from base filename  
## 9.2. Allowed filename tags  
- `[defaultNew]` (signature template files only)  
    - Set signature as default signature for new mails  
- `[defaultReplyFwd]` (signature template files only)  
    - Set signature as default signature for replies and forwarded mails  
- `[internal]` (OOF template files only)  
    - Set template as default OOF message for internal recipients  
    - If neither `[internal]` nor `[external]` is defined, the template is set as default OOF message for internal and external recipients  
- `[external]` (OOF template files only)  
    - Set template as default OOF message for external recipients  
    - If neither `[internal]` nor `[external]` is defined, the template is set as default OOF message for internal and external recipients  
- `[<NETBIOS Domain> <Group SamAccountName>]`, e.g. `[EXAMPLE Domain Users]`  
    - Make this template specific for an Outlook mailbox or the currently logged-on user being a member (direct or indirect) of this group  
    - Groups must be available in Active Directory. Groups like `'Everyone'` and `'Authenticated Users'` only exist locally, not in Active Directory  
- `[<SMTP address>]`, e.g. `[office<area>@example.com]`  
    - Make this template specific for the assigned mail address (all SMTP addresses of a mailbox are considered, not only the primary one)  
- `[yyyyMMddHHmm-yyyyMMddHHmm]`, e.g. `[202112150000-202112262359]` for the 2021 Christmas season  
    - Make this template valid only during the specific time range (`yyyy` = year, `MM` = month, `dd` = day, `HH` = hour, `mm` = minute)  
    - If the script does not run after a template has expired, the template is still available on the client and can be used.

Filename tags can be combined: A template may be assigned to several groups, several mail addresses and several time ranges, be used as default signature for new e-mails and as default signature for replies and forwards at the same time.

The number of possible tags is limited by Operating System file name and path length restrictions only. The script works with path names longer than the default Windows limit of 260 characters, even when "LongPathsEnabled" (https://docs.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation) is not active.  
# 10. Signature and OOF application order  
Templates are applied in a specific order: Common tempaltes first, group templates second, mail address specific templates last.

Templates with a time range tag are only considered if the current system time is in range of at least one of these tags.

Common templates are templates with either no tag or only `[defaultNew]` and/or `[defaultReplyFwd]` (`[internal]` and/or `[external]` for OOF templates).

Within these groups, templates are applied alphabetically ascending.

Every centrally stored signature template is applied only once, as there is only one signature path in Outlook, and subfolders are not allowed - so the file names have to be unique.

The script always starts with the mailboxes in the default Outlook profile, preferrably with the current users personal mailbox.

OOF templates are only applied if the Out of Office assistant is currently disabled. If it is currently active or scheduled to be activated in the future, OOF templates are not applied.  
# 11. Variable replacement  
Variables are case sensitive.

Variables are replaced everywhere, including links, QuickTips and alternative text of images.

With this feature, you can not only show mail addresses and telephone numbers in the signature and OOF message, but show them as links which open a new mail message (`"mailto:"`) or dial the number (`"tel:"`) via a locally installed softphone when clicked.

Custom Active directory attributes are supported as well as custom replacement variables, see `'.\config\default replacement variables.ps1'` for details.

Variables can also be retrieved from other sources than Active Directory by adding custom code to the variable config file.

Per default, `'.\config\default replacement variables.ps1'` contains the following replacement variables:  
- Currently logged-on user  
    - `$CURRENTUSERGIVENNAME$`: Given name  
    - `$CURRENTUSERSURNAME$`: Surname  
    - `$CURRENTUSERDEPARTMENT$`: Department  
    - `$CURRENTUSERTITLE$`: Title  
    - `$CURRENTUSERSTREETADDRESS$`: Street address  
    - `$CURRENTUSERPOSTALCODE$`: Postal code  
    - `$CURRENTUSERLOCATION$`: Location  
    - `$CURRENTUSERCOUNTRY$`: Country  
    - `$CURRENTUSERTELEPHONE$`: Telephone number  
    - `$CURRENTUSERFAX$`: Facsimile number  
    - `$CURRENTUSERMOBILE$`: Mobile phone  
    - `$CURRENTUSERMAIL$`: Mail address  
    - `$CURRENTUSERPHOTO$`: Photo from Active Directory, see "[11.1 Photos from Active Directory](#111-photos-from-active-directory)" for details  
    - `$CURRENTUSERPHOTODELETEEMPTY$`: Photo from Active Directory, see "[11.1 Photos from Active Directory](#111-photos-from-active-directory)" for details  
    - `$CURRENTUSEREXTATTR1$` to `$CURRENTUSEREXTATTR15$`: Exchange Extension Attributes 1 to 15  
- Manager of currently logged-on user  
    - Same variables as logged-on user, `$CURRENTUSERMANAGER\[...]$` instead of `$CURRENTUSER\[...]$`  
- Current mailbox  
    - Same variables as logged-on user, `$CURRENTMAILBOX\[...]$` instead of `$CURRENTUSER\[...]$`  
- Manager of current mailbox  
    - Same variables as logged-on user, `$CURRENTMAILBOXMANAGER\[...]$` instead of `$CURRENTMAILBOX[...]$`  
## 11.1. Photos from Active Directory  
The script supports replacing images in signature templates with photos stored in Active Directory.

When using images in OOF templates, please be aware that Exchange and Outlook do not yet support images in OOF messages.

As with other variables, photos can be obtained from the currently logged-on user, it's manager, the currently processed mailbox and it's manager.
  
To be able to apply Word image features such as sizing, cropping, frames, 3D effects etc, you have to exactly follow these steps:  
1. Create a sample image file which will later be used as placeholder.  
2. Optionally: If the sample image file name contains one of the following variable names, the script recognizes it and you do not need to add the value to the alternative text of the image in step 4:  
    - `$CURRENTUSERPHOTO$`  
    - `$CURRENTUSERPHOTODELETEEMPTY$`  
    - `$CURRENTUSERMANAGERPHOTO$`  
    - `$CURRENTUSERMANAGERPHOTODELETEEMPTY$`  
    - `$CURRENTMAILBOXPHOTO$`  
    - `$CURRENTMAILBOXPHOTODELETEEMPTY$`  
    - `$CURRENTMAILBOXMANAGERPHOTO$`  
    - `$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$`  
3. Insert the image into the signature template. Make sure to use `Insert | Pictures | This device` (Word 2019, other versions have the same feature in different menus) and to select the option `Insert and Link` - if you forget this step, a specific Word property is not set and the script will not be able to replace the image.  
4. If you did not follow optional step 2, please add one of the following variable names to the alternative text of the image in Word (these variables are removed from the alternative text in the final signature):  
    - `$CURRENTUSERPHOTO$`  
    - `$CURRENTUSERPHOTODELETEEMPTY$`  
    - `$CURRENTUSERMANAGERPHOTO$`  
    - `$CURRENTUSERMANAGERPHOTODELETEEMPTY$`  
    - `$CURRENTMAILBOXPHOTO$`  
    - `$CURRENTMAILBOXPHOTODELETEEMPTY$`  
    - `$CURRENTMAILBOXMANAGERPHOTO$`  
    - `$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$`  
5. Format the image as wanted.

For the script to recognize images to replace, you need to follow at least one of the steps 2 and 4. If you follow both, the script first checks for step 2 first. If you provide multiple image replacement variables, `$CURRENTUSER[...]$` has the highest priority, followed by `$CURRENTUSERMANAGER[...]$`, `$CURRENTMAILBOX[...]$` and `$CURRENTMAILBOXMANAGER[...]$`. It is recommended to use only one image replacement variable per image.  
  
The script will replace all images meeting the conditions described in the steps above and replace them with Active Directory photos in the background. This keeps Word image formatting option alive, just as if you would use Word's `"Change picture"` function.  
  
If there is no photo available in Active Directory, there are two options:  
- You used the `$CURRENT[...]PHOTO$` variables: The sample image used as placeholder is shown in the signature.  
- You used the `$CURRENT[...]PHOTODELETEEMPTY$` variables: The sample image used as placeholder is deleted from the signature, which may affect the layout of the remaining signature depending on your formatting options.

**Attention**: A signature with embedded images has the expected file size in DOCX, HTML and TXT formats, but the RTF file will be much bigger.

The signature template `'.\templates\Signatures DOCX\Test all signature replacement variables.docx'` contains several embedded images and can be used for a file comparison:  
- .docx: 23 KB  
- .htm: 87 KB  
- .RTF without workaround: 27.5 MB  
- .RTF with workaround: 1.4 MB
  
The script uses a workaround, but the resulting RTF files are still huge compared to other file types and especially for use in emails. If this is a problem, please either do not use embedded images in the signature template (including photos from Active Directory), or switch to HTML formatted emails.

If you ran into this problem outside this script, please consider modifying the ExportPictureWithMetafile setting as described in https://support.microsoft.com/kb/224663. If the link is not working, please visit the Internet Archive Wayback Machine's snapshot of Microsoft's article at https://web.archive.org/web/20180827213151/https://support.microsoft.com/en-us/help/224663/document-file-size-increases-with-emf-png-gif-or-jpeg-graphics-in-word.  
# 12. Outlook Web  
If the currently logged-on user has configured his personal mailbox in Outlook, the default signature for new emails is configured in Outlook Web automatically.

If the default signature for new mails matches the one used for replies and forwarded mail, this is also set in Outlook.

If different signatures for new and reply/forward are set, only the new signature is copied to Outlook Web.

If only a default signature for replies and forwards is set, only this new signature is copied to Outlook Web.

If there is no default signature in Outlook, Outlook Web settings are not changed.

All this happens with the credentials of the currently logged-on user, without any interaction neccessary.  
# 13. Simulation mode  
Simulation mode is enabled when the parameter SimulatedUser is passed to the script. It answers the question `"What will the signatures look like for user A, when Outlook is configured for the mailboxes X, Y and Z?"`.

Simulation mode is useful for content creators and admins, as it allows to simulate the behavior of the script and to inspect the resulting signature files before going live.
  
In simulation mode, Outlook registry entries are not considered and nothing is changed in Outlook and Outlook web. The template files are handled just as during a real script run, but only saved to the folder passed by the parameters AdditionalSignaturePath and AdditionalSignaturePath folder.
  
`SimulationUser` is a mandatory parameter for simulation mode. This value replaces the currently logged-on user.

`SimulationMailboxes` is optional for simulation mode, although highly recommended. It is a comma separated list of strings replacing the list of mailboxes otherwise gathered from the registry.

Active Directory data for both parameters is searched using Active Directory Ambigous Name Resolution (ANR), so you can pass very different values to find the desired object (mail address, logon name, display name, etc.). Please see https://social.technet.microsoft.com/wiki/contents/articles/22653.active-directory-ambiguous-name-resolution.aspx for details about ANR.

**Attention**:  
- Use values that are unique in an Active Directoy forest, not just in a domain. The script queries against the Global Catalog and always works with the first result returned only (even if there are additional results). For example, the logon name (sAMAccountName) must be unique within an Active Directory domain, but each domain in an Active Directory forest can have one account with this logon name. The script informs when there is more than one or no result.  
- Simulation mode only works when the user starting the simulation is at least from the same Active Directory forest as the user defined in SimulationUser.  Users from other forests will not work.  
# 14. FAQ
## 14.1. Where can I find the changelog?
The changelog is located in the `'.\docs'` folder, along with other documents related to Set-OutlookSignatures.
## 14.2. How can I contribute, propose a new feature or file a bug?
If you have an idea for a new feature or have found a problem, please [create an issue on GitHub](https://github.com/GruberMarkus/Set-OutlookSignatures/issues).

If you want to contribute code, please have a look at `'.\docs\CONTRIBUTING'` for a rough overview of the proposed process.
## 14.3. Why use legacyExchangeDN to find the user behind a mailbox, and not mail or proxyAddresses?  
The legacyExchangeDN attribute is used to find the user behind a mailbox, because mail and proxyAddresses are not unique in certain Exchange scenarios:  
- A separate Active Directory forest for users and Exchange mailboxes: In this case, the mail attribute is usually set in the user forest, although there are no mailboxes in this forest.  
- One common mail domain across multiple Exchange organizations: In this case, the address book is very like synchronized between Active Directory forests by using contacts or mail-enabled users, which both will have the SMTP address of the mailbox in the proxyAddresses attribute.

The disadvantage of using legacyExchangeDN is that no group membership information can be retrieved for Exchange mailboxes configured as IMAP or POP accounts in Outlook. This scenario is very rare in Exchange/Outlook enterprise environments. These mailboxes can still receive common and mailbox specific signatures and OOF messages.  
## 14.4. How is the personal mailbox of the currently logged-on user identified?  
The personal mailbox of the currently logged-on user is preferred to other mailboxes, as it receives signatures first and is the only mailbox where the Outlook Web signature can be set.

The personal mailbox is found by simply checking if the Active Directory mail attribute of the currently logged-on user matches an SMTP address of one of the mailboxes connected in Outlook.

If the mail attribute is not set, the currently logged-on user's objectSID is compared with all the mailboxes' msExchMasterAccountSID. If there is exactly one match, this mailbox is used as primary one.
  
Please consider the following caveats regarding the mail attribute:  
- When Active Directory attributes are directly modified to create or modify users and mailboxes (instead of using Exchange Admin Center or Exchange Management Shell), the mail attribute is often not updated and does not match the primary SMTP address of a mailbox. Microsoft strongly recommends that the mail attribute matches the primary SMTP address.  
- When using linked mailboxes, the mail attribute of the linked account is often not set or synced back from the Exchange resource forest. Technically, this is not necessary. From an organizational point of view it makes sense, as this can be used to determine if a specific user has a linked mailbox in another forest, and as some applications (such as "scan to mail") may need this attribute anyhow.  
## 14.5. Which ports are required?  
Ports 389 (LDAP) and 3268 (Global Catalog), both TCP and UDP, are required to communicate with Active Directory domains.

The client needs the following ports to access a SMB file share on a Windows server: 137 UDP, 138 UDP, 139 TCP, 445 TCP (for details, see https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc731402(v=ws.11).

The client needs port 443 to access a WebDAV share (a SharePoint document library, for example).  
## 14.6. Why is Out of Office abbreviated OOF and not OOO?  
Back in the 1980s, Microsoft had a UNIX OS named Xenix ... but read yourself: https://techcommunity.microsoft.com/t5/exchange-team-blog/why-is-oof-an-oof-and-not-an-ooo/ba-p/610191  
## 14.7. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.  
The script uses DOCX as default template format, as this seems to be the easiest way to delegate the creation and management of templates to departments such as Marketing or Corporate Communications:  
- Not all Word formatting options are supported in HTML, which can lead to signatures looking a bit different than templates. For example, images may be placed at a different position in the signature compared to the template - this is because the Outlook HTML component only supports the "in line with text" text wrapping option, while Word offers more options.  
- On the other hand, the Outlook HTML renderer works better with templates in the DOCX format: The Outlook HTML renderer does not respect the HTML image tags "width" and "height" and displays all images in their original size. When using DOCX as template format, the images are resized when exported to the HTM format.
  
I recommend to start with .docx as template format and to only use .htm when the template maintainers have really good HTML knowledge.

With the parameter `UseHtmTemplates`, the script searches for .htm template files instead of DOCX.

The requirements for .htm files these files are harder to fulfill as it is the case with DOCX files:  
- The template must be UTF8 encoded, or at least only contain UTF8 compatible characters  
- The template should be a single file, additional files and folders are not recommended  
- Images should ideally either reference a public URL or be part of the template as Base64 encoded string  
- The template must have the file extension .htm, .html is not supported
  
Possible approaches for fulfilling these requirements are:  
- Design the template in a HTML editor that supports all features required  
- Design the template in Outlook  
  - Paste it into Word and save it as `"Website, filtered"`. The `"filtered"` is important here, as any other web format will not work.  
  - Run the resulting file through a script that converts the Word output to a single UTF8 encoded HTML file. Alternatively, but not recommended, you can copy the .htm file and the associated folder containing images and other HTML information into the template folder.

You can use the script function ConvertTo-SingleFileHTML for embedding:
```
get-childitem ".\templates\Signatures HTML" -File | foreach-object {
    $_.FullName  
    ConvertTo-SingleFileHTML $_.FullName ($_.FullName -replace ".htm$", " embedded.htm")
} 
```

The templates delivered with this script represent all possible formats:  
- `'.\templates\Out of Office DOCX'` and `'.\templates\signatures DOCX'` contain templates in the DOCX format  
- `'.\templates\Out of Office HTML'` contains templates in the HTML format as Word exports them when using `"Website, filtered"` as format. Note the additional folders for each signature.  
- `'.\templates\Signatures HTML'` contains templates in the HTML format. Note that there are no additional folders, as the Word export files have been processed with ConvertTo-SingleFileHTML function to create a single HTMl file with all local images embedded.  
## 14.8. How can I log the script output?  
The script has no built-in logging option other than writing output to the host window.

You can, for example, use PowerShell's `Start-Transcript` and `Stop-Transcript` commands to create a logging wrapper around Set-OutlookSignatures.ps1.  
## 14.9. Can multiple script instances run in parallel?  
The script is designed for being run in multiple instances at the same. You can combine any of the following scenarios:  
- One user runs multiple instances of the script in parallel  
- One user runs multiple instances of the script in simulation mode in parallel  
- Multiple users on the same machine (e.g. Terminal Server) run multiple instances of the script in parallel  
## 14.10. How do I start the script from the command line or a scheduled task?  
Passing arguments to PowerShell.exe from the command line or task scheduler can be very tricky when spaces are involved. You have to be very careful about when to use single quotes or double quotes.

A working example:
```
PowerShell.exe -Command "& '\\server\share\directory\Set-OutlookSignatures.ps1' -SignatureTemplatePath '\\server\share\directory\templates\Signatures DOCX' -OOFTemplatePath '\\server\share\directory\templates\Out of Office DOCX' -ReplacementVariableConfigFile '\\server\share\directory\config\default replacement variables.ps1'"
```
You will find lots of information about this topic on the internet. The following links provide a first starting point:  
- https://stackoverflow.com/questions/45760457/how-can-i-run-a-powershell-script-with-white-spaces-in-the-path-from-the-command
- https://stackoverflow.com/questions/28311191/how-do-i-pass-in-a-string-with-spaces-into-powershell
- https://stackoverflow.com/questions/10542313/powershell-and-schtask-with-task-that-has-a-space
  
If you have to use the PowerShell.exe `-Command` or `-File` parameter depends on details of your configuration, for example AppLocker in combination with PowerShell. You may also want to consider the `-EncodedCommand` parameter to start Set-OutlookSignatures.ps1 and pass parameters to it.
  
If you provided your users a link so they can start Set-OutlookSignatures.ps1 with the correct parameters on their own, you may want to use the official icon: `'.\logo\Set-OutlookSignatures Icon.ico'`  
## 14.11. How to create a shortcut to the script with parameters?  
You may want to provide a link on the desktop or in the start menu, so they can start the script on their own.

The Windows user interface does not allow you to create a shortcut with a combined length of full target path and arguments greater than 259 characters.

You can overcome this user interface limitation by using PowerShell to create a shortcut (.lnk file):  
```
$WshShell = New-Object -ComObject WScript.Shell  
$Shortcut = $WshShell.CreateShortcut((Join-Path -Path $([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)) -ChildPath 'Set Outlook signatures.lnk'))  
$Shortcut.WorkingDirectory = '\\Long-Server-Name\Long-Share-Name\Long-Folder-Name\Set-OutlookSignatures'  
$Shortcut.TargetPath = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'  
$Shortcut.Arguments = "-NoExit -Command ""& '\\Long-Server-Name\Long-Share-Name\Long-Folder-Name\Set-OutlookSignatures\Set-OutlookSignatures.ps1' -SignatureTemplatePath '\\Long-Server-Name\Long-Share-Name\Long-Folder-Name\Templates\Signatures DOCX' -OOFTemplatePath '\\Long-Server-Name\Long-Share-Name\Long-Folder-Name\Templates\Out of Office DOCX'"""  
$Shortcut.IconLocation = '\\Long-Server-Name\Long-Share-Name\Long-Folder-Name\Set-OutlookSignatures\logo\Set-OutlookSignatures Icon.ico'  
$Shortcut.Description = 'Set Outlook signatures using Set-OutlookSignatures.ps1'  
$Shortcut.WindowStyle = 1 # 1 = undefined, 3 = maximized, 7 = minimized  
$Shortcut.Hotkey = ''  
$Shortcut.Save()  
```
**Attention**: When editing the shortcut created with the code above in the Windows user interface, the command to be executed is shortened to 259 characters without further notice. This already happens when just opening the properties of the created .lnk file, changing nothing and clicking OK.  
## 14.12. What is the recommended approach for implementing the software?  
There is certainly no definitive generic recommendation, but the file `'.\docs\Implementation approach.html'` should be a good starting point.

The content is based on real-life experience implementing the script in a multi-client environment with a five-digit number of mailboxes.

It contains proven procedures and recommendations for product managers, architects, operations managers, account managers and mail and client administrators. It is suited for service providers as well as for clients.

It covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The document is available in English and German language.  
## 14.13. What about the new signature roaming feature Microsoft announced?  
Microsoft announced a change in how and where signatures are stored. Basically, signatures are no longer stored in the file system, but in the mailbox itself.

This is a good idea, as it makes signatures available across devices and avoids file naming conflicts which may appear in current solutions.

Based on currently available information, the disadvantage is that signatures for shared mailboxes can no longer be personalized, as the latest signature change would be propagated to all users accessing the shared mailbox (which is especially bad when personalized signatures for shared mailboxes are set as default signature).

Microsoft has stated that only cloud mailboxes support the new feature and that Outlook for Windows will be the only client supporting the new feature for now. I am confident more mail clients will follow soon. Future will tell if the feature will be made available for mailboxes on premises, too.

Currently, there is no detailed documentation and no API available to programatically access the new feature.

Until the feature is fully rolled out and an API is available, you can disable the feature with a registry key. This forces Outlook for Windows to use the well-known file based approach and ensures full compatibility with this script.

For details, please see https://support.microsoft.com/en-us/office/outlook-roaming-signatures-420c2995-1f57-4291-9004-8f6f97c54d15?ui=en-us&rs=en-us&ad=us.  
