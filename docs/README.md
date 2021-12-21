<!-- omit in toc -->
# <a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="../src/logo/Set-OutlookSignatures%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-OutlookSignatures"></a><br>Centrally manage and deploy Outlook text signatures and Out of Office auto reply messages.<br><a href="https://github.com/GruberMarkus/Set-OutlookSignatures/blob/main/docs/LICENSE.txt" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Set-OutlookSignatures" alt=""></a> <a href="https://www.paypal.com/donate/?business=JBM584K3L5PX4&item_name=Set-OutlookSignatures&no_recurring=0&currency_code=EUR" target="_blank"><img src="https://img.shields.io/badge/sponsor-grey?logo=paypal" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Set-OutlookSignatures/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a>  

# Features <!-- omit in toc -->
**Signatures and OOF messages can be:**
- Generated from templates in DOCX or HTML file format  
- Customized with a broad range of variables, including photos, from Active Directory and other sources  
- Applied to all mailboxes (including shared mailboxes), specific mailbox groups or specific e-mail addresses, for every primary mailbox across all Outlook profiles  
- Assigned time ranges within which they are valid  
- Set as default signature for new mails, or for replies and forwards (signatures only)  
- Set as default OOF message for internal or external recipients (OOF messages only)  
- Set in Outlook Web for the currently logged in user  
- Centrally managed only or exist along user created signatures (signatures only)  
- Copied to an alternate path for easy access on mobile devices not directly supported by this script (signatures only)

Set-Outlooksignatures can be **executed by users on clients, or on a server without end user interaction**.  
On clients, it can run as part of the logon script, as scheduled task, or on user demand via a desktop icon, start menu entry, link or any other way of starting a program.  
Signatures and OOF messages can also be created and deployed centrally, without end user or client involvement.

**Sample templates** for signatures and OOF messages demonstrate all available features and are provided as .docx and .htm files.

**Simulation mode** allows content creators and admins to simulate the behavior of the script and to inspect the resulting signature files before going live.
  
The script is **designed to work in big and complex environments** (Exchange resource forest scenarios, across AD trusts, multi-level AD subdomains, many objects). It works **on premises, in hybrid and cloud-only environments**.

It is **multi-client capable** by using different template paths, configuration files and script parameters.

Set-OutlookSignature requires **no installation on servers or clients**. You only need a standard file share on a server, and PowerShell and Office. 

A **documented implementation approach**, based on real-life experience implementing the script in a multi-client environment with a five-digit number of mailboxes, contains proven procedures and recommendations for product managers, architects, operations managers, account managers and e-mail and client administrators.  
The implementatin approach is **suited for service providers as well as for clients**, and covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The script is **Free and Open-Source Software (FOSS)**. It is published under the MIT license which is approved, among others, by the Free Software Foundation (FSF) and the Open Source Initiative (OSI), and is compatible with the General Public License (GPL) v3. Please see `'.\docs\LICENSE.txt'` for copyright and MIT license details.
# Table of Contents <!-- omit in toc -->
- [1. Requirements](#1-requirements)
- [2. Parameters](#2-parameters)
  - [2.1. SignatureTemplatePath](#21-signaturetemplatepath)
  - [2.2. SignatureIniPath](#22-signatureinipath)
  - [2.3. ReplacementVariableConfigFile](#23-replacementvariableconfigfile)
  - [2.4. GraphConfigFile](#24-graphconfigfile)
  - [2.5. TrustedDomainsToCheckForGroups](#25-trusteddomainstocheckforgroups)
  - [2.6. DeleteUserCreatedSignatures](#26-deleteusercreatedsignatures)
  - [2.7. DeleteScriptCreatedSignaturesWithoutTemplate](#27-deletescriptcreatedsignatureswithouttemplate)
  - [2.8. SetCurrentUserOutlookWebSignature](#28-setcurrentuseroutlookwebsignature)
  - [2.9. SetCurrentUserOOFMessage](#29-setcurrentuseroofmessage)
  - [2.10. OOFTemplatePath](#210-ooftemplatepath)
  - [2.11. OOFIniPath](#211-oofinipath)
  - [2.12. AdditionalSignaturePath](#212-additionalsignaturepath)
  - [2.13. AdditionalSignaturePathFolder](#213-additionalsignaturepathfolder)
  - [2.14. UseHtmTemplates](#214-usehtmtemplates)
  - [2.15. SimulateUser](#215-simulateuser)
  - [2.16. SimulateMailboxes](#216-simulatemailboxes)
  - [2.17. GraphCredentialFile](#217-graphcredentialfile)
  - [2.18. GraphOnly](#218-graphonly)
  - [2.19. CreateRTFSignatures](#219-creatertfsignatures)
  - [2.20. CreateTXTSignatures](#220-createtxtsignatures)
- [3. Outlook signature path](#3-outlook-signature-path)
- [4. Mailboxes](#4-mailboxes)
- [5. Group membership](#5-group-membership)
- [6. Removing old signatures](#6-removing-old-signatures)
- [7. Error handling](#7-error-handling)
- [8. Run script while Outlook is running](#8-run-script-while-outlook-is-running)
- [9. Signature and OOF file format](#9-signature-and-oof-file-format)
  - [9.1. Signature and OOF file naming](#91-signature-and-oof-file-naming)
  - [9.2. Allowed tags](#92-allowed-tags)
  - [9.3. Tags in ini files instead in file names](#93-tags-in-ini-files-instead-in-file-names)
- [10. Signature and OOF application order](#10-signature-and-oof-application-order)
- [11. Variable replacement](#11-variable-replacement)
  - [11.1. Photos from Active Directory](#111-photos-from-active-directory)
- [12. Outlook Web](#12-outlook-web)
- [13. Hybrid and cloud-only support](#13-hybrid-and-cloud-only-support)
  - [13.1. Basic Configuration](#131-basic-configuration)
  - [13.2. Advanced Configuration](#132-advanced-configuration)
  - [13.3. Authentication](#133-authentication)
- [14. Simulation mode](#14-simulation-mode)
- [15. FAQ](#15-faq)
  - [15.1. Where can I find the changelog?](#151-where-can-i-find-the-changelog)
  - [15.2. How can I contribute, propose a new feature or file a bug?](#152-how-can-i-contribute-propose-a-new-feature-or-file-a-bug)
  - [15.3. Why use legacyExchangeDN to find the user behind a mailbox, and not mail or proxyAddresses?](#153-why-use-legacyexchangedn-to-find-the-user-behind-a-mailbox-and-not-mail-or-proxyaddresses)
  - [15.4. How is the personal mailbox of the currently logged in user identified?](#154-how-is-the-personal-mailbox-of-the-currently-logged-in-user-identified)
  - [15.5. Which ports are required?](#155-which-ports-are-required)
  - [15.6. Why is Out of Office abbreviated OOF and not OOO?](#156-why-is-out-of-office-abbreviated-oof-and-not-ooo)
  - [15.7. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.](#157-should-i-use-docx-or-htm-as-file-format-for-templates-signatures-in-outlook-sometimes-look-different-than-my-templates)
  - [15.8. How can I log the script output?](#158-how-can-i-log-the-script-output)
  - [15.9. Can multiple script instances run in parallel?](#159-can-multiple-script-instances-run-in-parallel)
  - [15.10. How do I start the script from the command line or a scheduled task?](#1510-how-do-i-start-the-script-from-the-command-line-or-a-scheduled-task)
  - [15.11. How to create a shortcut to the script with parameters?](#1511-how-to-create-a-shortcut-to-the-script-with-parameters)
  - [15.12. What is the recommended approach for implementing the software?](#1512-what-is-the-recommended-approach-for-implementing-the-software)
  - [15.13. What is the recommended approach for custom configuration files?](#1513-what-is-the-recommended-approach-for-custom-configuration-files)
  - [15.14. Isn't a plural noun in the script name against PowerShell best practices?](#1514-isnt-a-plural-noun-in-the-script-name-against-powershell-best-practices)
  - [15.15. The script hangs at HTM/RTF export, Word shows a security warning!?](#1515-the-script-hangs-at-htmrtf-export-word-shows-a-security-warning)
  - [15.16. How to avoid empty lines when replacement variables return an empty string?](#1516-how-to-avoid-empty-lines-when-replacement-variables-return-an-empty-string)
  - [15.17. Is there a roadmap for future versions?](#1517-is-there-a-roadmap-for-future-versions)
  - [15.18. How to deploy signatures for "Send As", "Send On Behalf" etc.?](#1518-how-to-deploy-signatures-for-send-as-send-on-behalf-etc)
  - [15.19. Can I centrally manage and deploy Outook stationery with this script?](#1519-can-i-centrally-manage-and-deploy-outook-stationery-with-this-script)
  - [15.20. Why is membership in dynamic distribution groups and dynamic security groups not considered?](#1520-why-is-membership-in-dynamic-distribution-groups-and-dynamic-security-groups-not-considered)
    - [15.20.1. What's the alternative to dynamic groups?](#15201-whats-the-alternative-to-dynamic-groups)
  - [15.21. What about the new signature roaming feature Microsoft announced?](#1521-what-about-the-new-signature-roaming-feature-microsoft-announced)
    - [15.21.1. Please be aware of the following problem](#15211-please-be-aware-of-the-following-problem)
  
# 1. Requirements  
Requires Outlook and Word, at least version 2010.  
The script must run in the security context of the currently logged in user.

The script must run in PowerShell Full Language mode. Constrained Language mode is not supported, as some features such as BASE64 conversions are not available in this mode or require very slow workarounds.

If you use AppLocker or a comparable solution, you may need to digitally sign the PowerShell 'Set-OutlokSignatures.ps1'. It is usually not necessary to sign the variable replacement configuration files, e. g. '.\config\default replacement variables.ps1'.  
There are locked down environments, where all files matching the patterns "\*.ps\*1" and "*.dll" need to be digitially signed with a trusted certificate. 

Don't forget to unblock at least 'Set-OutlookSignatures.ps1' after extracting them from the downloaded ZIP file. You can use the PowerShell commandlet 'Unblock-File' for this.

The paths to the template files (SignatureTemplatePath, OOFTemplatePath) must be accessible by the currently logged in user. The template files must be at least readable for the currently logged in user.

In cloud environments, you need to register Set-OutlookSignatures as app and provide admin consent for the required permissions. See '.\config\default graph config.ps1' for details.
# 2. Parameters  
## 2.1. SignatureTemplatePath  
The parameter SignatureTemplatePath tells the script where signature template files are stored.

Local and remote paths are supported. Local paths can be absolute (`'C:\Signature templates'`) or relative to the script path (`'.\templates\Signatures'`).

WebDAV paths are supported (https only): `'https://server.domain/SignatureSite/SignatureTemplates'` or `'\\server.domain@SSL\SignatureSite\SignatureTemplates'`

The currently logged in user needs at least read access to the path.

Default value: `'.\templates\Signatures DOCX'`  
## 2.2. SignatureIniPath
If you can't or don't want to use file name based tags, you can place them in an ini file.

See '.\templates\sample signatures ini file.ini' for a sample file with further explanations.

Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures')

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'

The currently logged in user needs at least read access to the path

Default value: `''`
## 2.3. ReplacementVariableConfigFile  
The parameter ReplacementVariableConfigFile tells the script where the file defining replacement variables is located.

Local and remote paths are supported. Local paths can be absolute (`'C:\config\default replacement variables.ps1'`) or relative to the script path (`'.\config\default replacement variables.ps1'`).

WebDAV paths are supported (https only): `'https://server.domain/SignatureSite/config/default replacement variables.ps1'` or `'\\server.domain@SSL\SignatureSite\config\default replacement variables.ps1'`

The currently logged in user needs at least read access to the file.

Default value: `'.\config\default replacement variables.ps1'`  
## 2.4. GraphConfigFile
The parameter GraphConfigFile tells the script where the file defining Graph connection and configuration options is located.

Local and remote paths are supported. Local paths can be absolute (`'C:\config\default graph config.ps1'`) or relative to the script path (`'.\config\default graph config.ps1'`).

WebDAV paths are supported (https only): `'https://server.domain/SignatureSite/config/default graph config.ps1'` or `'\\server.domain@SSL\SignatureSite\config\default graph config.ps1'`

The currently logged in user needs at least read access to the file.

Default value: `'.\config\default graph config.ps1'`  
## 2.5. TrustedDomainsToCheckForGroups  
The parameters tells the script which trusted domains should be used to search for mailbox and user group membership.

The default value, `'*'` tells the script to query all trusted domains in the Active Directory forest of the logged in user.

For a custom list of trusted domains, specify them as comma-separated list of strings: `"domain-a.local", "dc=example,dc=com", "domain-b.internal"`.

When a domain in the custom list starts with a dash or minus (`'-domain-a.local'`), this domain is removed from the list.

The `'*'` entry in a custom list is only considered when it is the first entry of the list.

The Active Directory forest of the currently logged in user is always considered.

Subdomains of trusted domains are always considered.

Default value: `'*'`  
## 2.6. DeleteUserCreatedSignatures  
Shall the script delete signatures which were created by the user itself?

Default value: `$false`
## 2.7. DeleteScriptCreatedSignaturesWithoutTemplate
Shall the script delete signatures which were created by the script before but are no longer available as template?

Default value: `$true`
## 2.8. SetCurrentUserOutlookWebSignature  
Shall the script set the Outlook Web signature of the currently logged in user?

If the parameter is set to `$true` and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. This way, the script can be used in environments where only Outlook Web is used. 

Default value: `$true`  
## 2.9. SetCurrentUserOOFMessage  
Shall the script set the Out of Office (OOF) auto reply message of the currently logged in user?

If the parameter is set to `$true` and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. This way, the script can be used in environments where only Outlook Web is used. 

Default value: `$true`  
## 2.10. OOFTemplatePath  
Path to centrally managed Out of Office (OOF) auto reply templates.

Local and remote paths are supported.

Local paths can be absolute (`'C:\OOF templates'`) or relative to the script path (`'.\templates\Out of Office'`).

WebDAV paths are supported (https only): `'https://server.domain/SignatureSite/OOFTemplates'` or `'\\server.domain@SSL\SignatureSite\OOFTemplates'`

The currently logged in user needs at least read access to the path.

Default value: `'.\templates\Out of Office DOCX'`
## 2.11. OOFIniPath
If you can't or don't want to use file name based tags, you can place them in an ini file.

See '.\templates\sample OOF ini file.ini' for a sample file with further explanations.

Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures')

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'

The currently logged in user needs at least read access to the path

Default value: `''`
## 2.12. AdditionalSignaturePath  
An additional path that the signatures shall be copied to.  
Ideally, this path is available on all devices of the user, for example via Microsoft OneDrive or Nextcloud.

This way, the user can easily copy-paste his preferred preconfigured signature for use in an e-mail app not supported by this script, such as Microsoft Outlook Mobile, Apple Mail, Google Gmail or Samsung Email.

Local and remote paths are supported.

Local paths can be absolute (`'C:\Outlook signatures'`) or relative to the script path (`'.\Outlook signatures'`).

WebDAV paths are supported (https only): `'https://server.domain/User/Outlook signatures'` or `'\\server.domain@SSL\User\Outlook signatures'`

The currently logged in user needs at least write access to the path.

If the folder or folder structure does not exist, it is created.

Default value: `"$([environment]::GetFolderPath("MyDocuments"))\Outlook signatures"`  
## 2.13. AdditionalSignaturePathFolder
A folder or folder structure below AdditionalSignaturePath.

This parameter is available for compatibility with versions before 2.2.1. Starting with 2.2.1, you can pass a full path via the parameter AdditionalSignaturePath, so AdditionalSignaturePathFolder is no longer needed.

If the folder or folder structure does not exist, it is created.

Default value: `'Outlook signatures'`  
## 2.14. UseHtmTemplates  
With this parameter, the script searches for templates with the extension .htm instead of .docx.

Each format has advantages and disadvantages, please see "[13.5. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.](#135-should-i-use-docx-or-htm-as-file-format-for-templates-signatures-in-outlook-sometimes-look-different-than-my-templates)" for a quick overview.

Default value: `$false`  
## 2.15. SimulateUser  
SimulateUser is a mandatory parameter for simulation mode. This value replaces the currently logged in user.

Use a logon name in the format 'Domain\User' or a Universal Principal Name (UPN, looks like an e-mail-address, but is not neecessarily one).

See "[13. Simulation mode](#13-simulation-mode)" for details.  
## 2.16. SimulateMailboxes  
SimulateMailboxes is optional for simulation mode, although highly recommended. It is a comma separated list of e-mail addresses replacing the list of mailboxes otherwise gathered from the registry.
## 2.17. GraphCredentialFile
Path to file containing Graph credential which should be used as alternative to other token acquisition methods.

Makes only sense in combination with `'.\sample code\SimulateAndDeploy.ps1'`, do not use this parameter for other scenarios.

See `'.\sample code\SimulateAndDeploy.ps1'` for an example how to create this file.

Default value: `$null`  
## 2.18. GraphOnly
Try to connect to Microsoft Graph only, ignoring any local Active Directory.

The default behavior is to try Active Directory first and fall back to Graph.

Default value: `$false`
## 2.19. CreateRTFSignatures
Should signatures be created in RTF format?

Default value: `$true`
## 2.20. CreateTXTSignatures
Should signatures be created in TXT format?

Default value: `$true`
# 3. Outlook signature path  
The Outlook signature path is retrieved from the users registry, so the script is language independent.

The registry setting does not allow for absolute paths, only for paths relative to `'%APPDATA%\Microsoft'`.

If the relative path set in the registry would be a valid path but does not exist, the script creates it.  
# 4. Mailboxes  
The script only considers primary mailboxes, these are mailboxes added as separate accounts.

This is the same way Outlook handles mailboxes from a signature perspective: Outlook can not handle signatures for non-primary mailboxes (added via "Open these additional mailboxes").

The script is created for Exchange environments. Non-Exchange mailboxes can not have OOF messages or group signatures, but common and mailbox specific signatures.  
# 5. Group membership  
The script considers all static security and distribution groups the currently processed mailbox belongs to.

Group membership is evaluated against the whole Active Directory forest of the currently logged in user, and against all trusted domains (and their subdomains) the user has access to.

In Exchange resource forest scenarios with linked mailboxes, the group membership of the linked account (as populated in msExchMasterAccountSID) is not considered, only the group membership of the actual mailbox.

Group membership from Active Directory on-prem is retrieved by combining two queries:
- Security groups, no matter if enabled for e-mail or not, are queried via the tokenGroups attribute. Querying this attribute is very fast, resource saving on client and server, and also considers sIDHistory.
- Distribution groups are not covered by the tokenGroups attribute, they are retrieved with an optimized LDAP query, also sIDHistory is considered.
- Group membership in any type of group across trusts is retrieved with an optimized LDAP query, considering the sID and sIDHistory of the group memberships retrieved in the steps before.

When no Active Directory connection is available, Microsoft Graph is queried for transitive group membership. This query includes security and distribution groups.

Only static groups are considered. Please see the FAQ section for detailed information why dynamic groups are not included in group membership queries.
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

If you can't or don't want to use file name based tags, you can also place them in an ini file.  
See the '.\templates' folder for sample templates and configuration.  
See the `'Tags in ini files instead in file names'` section for more details.

Examples:  
- `'Company external German.docx'` -> `'Company external German.htm'`, no changes  
- `'Company external German.[defaultNew].docx'` -> `'Company external German.htm'`, tag(s) is/are removed  
- `'Company external [English].docx'` -> `'Company external [English].htm'`, tag(s) is/are not removed, because there is no dot before  
- `'Company external [English].[defaultNew] [Company-AD All Employees].docx'` -> `'Company external [English].htm'`, tag(s) is/are removed, because they are separated from base filename  
## 9.2. Allowed tags  
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
    - Make this template specific for an Outlook mailbox or the currently logged in user being a member (direct or indirect) of this group  
    - Groups must be available in Active Directory. Groups like `'Everyone'` and `'Authenticated Users'` only exist locally, not in Active Directory

    This tag supports alternative formats, which are of special interest if you are in a cloud only or hybrid environmonent:
    -  `[<NETBIOS Domain> <Group SamAccountName>]` and `[<NETBIOS Domain> <Group DisplayName>]` can be queried from Microsoft Graph if the groups are synced between on-prem and the cloud. SamAccountName is queried before DisplayName.  
    Use these formats when your environment is hybrid or on premises only.
    -  `[AzureAD <Group e-mail address>]`, `[AzureAD <Group MailNickname>]`, `[AzureAD <Group DisplayName>]` do not work with a local Active Directory. They are queried in the order given.  
    'AzureAD' is the literal, case-insensitive string 'AzureAD', not a variable.  
    Use these formats when you are in a cloud only environment.  
  
  When using an ini file instead of filename based tags, you can negate a group by prefixing it with '-:'. This deny removes previously included mailboxes. Denies are stronger than allows, no matter in which order they appear within a template section in the ini file.  
  Denies are available for all kinds of templates: Time based, common, group specific and e-mail address specific.  
  Example:
  ```
  [OOF template.docx]
  # Valid for all mailboxes being direct or indirect members of "DOMAIN\Group", but not if they are direct or indirect members of "DOMAIN\OtherGroup" or if the mailbox has the e-mail address x@example.com
  DOMAIN Group
  -:DOMAIN OtherGroup
  -:x@example.com
  ```
- `[<SMTP address>]`, e.g. `[office<area>@example.com]`  
    - Make this template specific for the assigned e-mail address (all SMTP addresses of a mailbox are considered, not only the primary one)  
  
  When using an ini file instead of filename based tags, you can negate an e-mail address by prefixing it with '-:'. This deny removes previously included mailboxes. Denies are stronger than allows, no matter in which order they appear within a template section in the ini file.  
  Denies are available for all kinds of templates: Time based, common, group specific and e-mail address specific.  
  Example:
  ```
  [Signature template.docx]
  # Valid for the mailboxes with the SMTP address x@example.com and y@example.com, but not if they are direct or indirect members of "DOMAIN\OtherGroup"
  x@example.com
  y@example.com
  -:DOMAIN OtherGroup 
  ```
- `[yyyyMMddHHmm-yyyyMMddHHmm]`, e.g. `[202112150000-202112262359]` for the 2021 Christmas season  
    - Make this template valid only during the specific time range (`yyyy` = year, `MM` = month, `dd` = day, `HH` = hour, `mm` = minute)  
    - If the script does not run after a template has expired, the template is still available on the client and can be used.

Filename tags can be combined: A template may be assigned to several groups, several e-mail addresses and several time ranges, be used as default signature for new e-mails and as default signature for replies and forwards at the same time.

The number of possible tags is limited by Operating System file name and path length restrictions only.  
On Powershell 7+, the script works with path names longer than the default Windows limit of 260 characters.  
On Powershell 5.1, enable "LongPathsEnabled" on the Operating System level as described in <a href="https://docs.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation" target="_blank">this Microsoft article</a>.  
## 9.3. Tags in ini files instead in file names
Using an ini file has the following advantages:
- shorter template file names, as tags are in the ini file and no longer in the file names
- unlimited number of tags, as no file system restrictions apply
- different configurations for the same templates folder by using different ini files for different audiences
- alternative sort orders for templates within template groups (common, group specific, e-mail address specific)
- with file name tags, the application order is always alphabetically ascending using the system culture sort order - with ini files, you can switch to alphabetically descending or as sorted in the ini file and define an other sort culture

If you want to give template creators control over the ini file, place it in the same folder as the templates.

How to work with ini files:
1. Comments
  Comment lines start with '#' or ';'
	Whitespace(s) at the beginning and the end of a line are ignored
  Empty lines are ignored
2. Use the ini files in `'.\templates\Signatures DOCX with ini'` and `'.\templates\Out of Office DOCX with ini'` as templates and starting point
3. Put file names with extensions in square brackets  
  Example: `[Company external English formal.docx]`  
  Putting file names in single or double quotes is possible, but not necessary.  
  File names are case insensitive
    `[file a.docx]` is the same as `["File A.docx"]` and `['fILE a.dOCX']`  
  When there are two or more sections for a filename: The keys and values are not combined, only the last section is considered.  
  File names not mentioned in this file are not considered, even if they are available in the file system.
2. Add tags in the lines below the filename
  Example: `defaultNew`  
  - Do not enclose tags in square brackets. This is not allowed here, but required when you add tags directly to file names.  
  - When an ini file is used, tags in file names are not considered as tags, but as part of the file name, so the Outlook signature name will contain them.  
  - Only one tag per line is allowed.  
  Adding not a single tag to file name section is valid. The signature template is then classified as a common template.
  - Putting file names in single or double quotes is possible, but not necessary
  - Tags are case insensitive  
    `defaultNew` is the same as `DefaultNew` and `dEFAULTnEW`
  - You can override the automatic Outlook signature name generation by setting OutlookSignatureName, e. g. `OutlookSignatureName = This is a custom signature name`  
  With this option, you can have different template file names for the same Outlook signature name. Search for "Marketing external English formal" in this file for examples. Take care of signature group priorities (common, group, e-mail address) and SortOrder parameter.
3. Remove the tags from the file names in the file system  
Else, the file names in the ini file and the file system do not match, which will result in some templates not being applied.  
It is recommended to create a copy of your template folder for tests
4. Make the script use the ini file by passing the 'SignatureIniPath' and/or 'OOFIniPath' parameter
# 10. Signature and OOF application order  
Templates are applied in a specific order: Common tempaltes first, group templates second, e-mail address specific templates last.

Templates with a time range tag are only considered if the current system time is in range of at least one of these tags.

Common templates are templates with either no tag or only `[defaultNew]` and/or `[defaultReplyFwd]` (`[internal]` and/or `[external]` for OOF templates).

Within these groups, templates are applied alphabetically ascending.

Every centrally stored signature template is applied only once, as there is only one signature path in Outlook, and subfolders are not allowed - so the file names have to be unique.

The script always starts with the mailboxes in the default Outlook profile, preferrably with the current users personal mailbox.

OOF templates are only applied if the Out of Office assistant is currently disabled. If it is currently active or scheduled to be activated in the future, OOF templates are not applied.  
# 11. Variable replacement  
Variables are case sensitive.

Variables are replaced everywhere, including links, QuickTips and alternative text of images.

With this feature, you can not only show e-mail addresses and telephone numbers in the signature and OOF message, but show them as links which open a new mail message (`"mailto:"`) or dial the number (`"tel:"`) via a locally installed softphone when clicked.

Custom Active directory attributes are supported as well as custom replacement variables, see `'.\config\default replacement variables.ps1'` for details.

Variables can also be retrieved from other sources than Active Directory by adding custom code to the variable config file.

Per default, `'.\config\default replacement variables.ps1'` contains the following replacement variables:  
- Currently logged in user  
    - `$CURRENTUSERGIVENNAME$`: Given name  
    - `$CURRENTUSERSURNAME$`: Surname  
    - `$CURRENTUSERDEPARTMENT$`: Department  
    - `$CURRENTUSERTITLE$`: (Job) Title  
    - `$CURRENTUSERSTREETADDRESS$`: Street address  
    - `$CURRENTUSERPOSTALCODE$`: Postal code  
    - `$CURRENTUSERLOCATION$`: Location  
    - `$CURRENTUSERCOUNTRY$`: Country  
    - `$CURRENTUSERTELEPHONE$`: Telephone number  
    - `$CURRENTUSERFAX$`: Facsimile number  
    - `$CURRENTUSERMOBILE$`: Mobile phone  
    - `$CURRENTUSERMAIL$`: E-mail address  
    - `$CURRENTUSERPHOTO$`: Photo from Active Directory, see "[11.1 Photos from Active Directory](#111-photos-from-active-directory)" for details  
    - `$CURRENTUSERPHOTODELETEEMPTY$`: Photo from Active Directory, see "[11.1 Photos from Active Directory](#111-photos-from-active-directory)" for details  
    - `$CURRENTUSEREXTATTR1$` to `$CURRENTUSEREXTATTR15$`: Exchange extension attributes 1 to 15  
    - `$CURRENTUSERCOMPANY`: Company  
    - `$CURRENTUSERMAILNICKNAME`: Alias (mailNickname)  
    - `$CURRENTUSERDISPLAYNAME`: Display Name  
- Manager of currently logged in user  
    - Same variables as logged in user, `$CURRENTUSERMANAGER[...]$` instead of `$CURRENTUSER[...]$`  
- Current mailbox  
    - Same variables as logged in user, `$CURRENTMAILBOX[...]$` instead of `$CURRENTUSER[...]$`  
- Manager of current mailbox  
    - Same variables as logged in user, `$CURRENTMAILBOXMANAGER[...]$` instead of `$CURRENTMAILBOX[...]$`  
## 11.1. Photos from Active Directory  
The script supports replacing images in signature templates with photos stored in Active Directory.

When using images in OOF templates, please be aware that Exchange and Outlook do not yet support images in OOF messages.

As with other variables, photos can be obtained from the currently logged in user, it's manager, the currently processed mailbox and it's manager.
  
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

If you ran into this problem outside this script, please consider modifying the ExportPictureWithMetafile setting as described in  <a href="https://support.microsoft.com/kb/224663" target="_blank">this Microsoft article</a>.  
If the link is not working, please visit the <a href="https://web.archive.org/web/20180827213151/https://support.microsoft.com/en-us/help/224663/document-file-size-increases-with-emf-png-gif-or-jpeg-graphics-in-word" target="_blank">Internet Archive Wayback Machine's snapshot of Microsoft's article</a>.  
# 12. Outlook Web  
If the currently logged in user has configured his personal mailbox in Outlook, the default signature for new emails is configured in Outlook Web automatically.

If the default signature for new mails matches the one used for replies and forwarded mail, this is also set in Outlook.

If different signatures for new and reply/forward are set, only the new signature is copied to Outlook Web.

If only a default signature for replies and forwards is set, only this new signature is copied to Outlook Web.

If there is no default signature in Outlook, Outlook Web settings are not changed.

All this happens with the credentials of the currently logged in user, without any interaction neccessary.  
# 13. Hybrid and cloud-only support
Set-OutlookSignatures supports three directory environments:
- Active Directory on premises. This requires direct connection to Active Directory Domain Controllers, which usually only works when you are connected to your company network.
- Hybrid. This environment consists of an Active Directory on premises, which is synced with Microsoft 365 Azure Active Directory in the cloud. If the script can't make a connection to your on-prem environment, it tries to get required data from the cloud via the Microsoft Graph API.
- Cloud-only. This environment has no Active Directory on premises, only Microsoft 365 with Azure Active Directory is used. If the script can't make a connection to your on-prem environment, it tries to get required data from the cloud via the Microsoft Graph API.
## 13.1. Basic Configuration
To allow communication between Microsoft Graph and Set-Outlooksignatures, both need to be configured for each other.

The easiest way is to once start Set-OutlookSignatures with a cloud administrator. The administrator then gets asked for admin consent for the correct permissions.  
If you don't want to use custom Graph attributes or other advanced configurations, no more configuration in Microsoft Graph or Set-OutlookSignatures is required.

If you prefer using own application IDs or need advanced configuration, follow these steps:  
- In Microsoft Graph, with an administrative account:
  - Create an application with a Client ID
  - Provide admin consent (pre-approval) for the following scopes (permissions):
    - 'https<area>://graph.microsoft.com/openid' for logging-on the use
    - 'https<area>://graph.microsoft.com/email' for reading the logged in user's mailbox properties
    - 'https<area>://graph.microsoft.com/profile' for reading the logged in user's properties
    - 'https<area>://graph.microsoft.com/user.read.all' for reading properties of other users (manager, additional mailboxes and their managers)
    - 'https<area>://graph.microsoft.com/group.read.all' for reading properties of all groups, required for templates restricted to groups
    - 'https<area>://graph.microsoft.com/mailboxsettings.readwrite' for updating the user's own mailbox settings (Out of Office auto reply messages)
    - 'https<area>://graph.microsoft.com/EWS.AccessAsUser.All' for updating the Outlook Web signature in the user's own mailbox
  - Set the Redirect URI to 'http<area>://localhost', configure for 'mobile and desktop applications'
  - Enable 'Allow public client flows' to make Windows Integrated Authentication (SSO) work for Azure AD joined devices
- In Set-OutlookSignature, use '.\config\default graph config.ps1' as a template for a custom Graph configuration file
  - Set '$GraphClientID' to the application ID created by the Graph administrator before
  - Use the 'GraphConfigFile' parameter to make the tool use the newly created Graph configuration file.
## 13.2. Advanced Configuration
The Graph configuration file allows for additional, advanced configuration:
- '$GraphEndpointVersion': The version of the Graph REST API to use
- '$GraphUserProperties': The properties to load for each graph user/mailbox. You can add custom attributes here.
- '$GraphUserAttributeMapping': Graph and Active Directory attributes are not named identically. Set-OutlookSignatures therefore uses a "virtual" account. Use this hashtable to define which Graph attribute name is assigned to which attribute of the virtual account.  
The virtual account is accessible as '\$ADPropsCurrentUser\[...\]' in '.\config\default replacement variables.ps1', and therefore has a direct impact on replacement variables.
## 13.3. Authentication
In hybrid and cloud-only scenarios, Set-OutlookSignatures automatically tries three stages of authentication.
1. Windows Integrated Authentication  
  This works in hybrid scenarios. The credentials of the currently logged in user are used to access Microsoft Graph without any further user interaction.
2. Silent authentication  
  If Windows Integrated Authentication fails, the User Principal Name of the currently logged in user is determined. If an existing cached cloud credential for this UPN is found, it is used for authentication with Microsoft Graph.  
  A default browser window with an "Authentication successful" message may open, it can be closed anytime.
3. User interaction  
  If the other authentication methods fail, the user is interactively asked for credentials. No custom components are used, only the official Microsoft 365 authentication site and the user's default browser. 
# 14. Simulation mode  
Simulation mode is enabled when the parameter SimulatedUser is passed to the script. It answers the question `"What will the signatures look like for user A, when Outlook is configured for the mailboxes X, Y and Z?"`.

Simulation mode is useful for content creators and admins, as it allows to simulate the behavior of the script and to inspect the resulting signature files before going live.
  
In simulation mode, Outlook registry entries are not considered and nothing is changed in Outlook and Outlook web.

The template files are handled just as during a real script run, but only saved to the folder passed by the parameters AdditionalSignaturePath and AdditionalSignaturePath folder.
  
`SimulateUser` is a mandatory parameter for simulation mode. This value replaces the currently logged in user. Use a logon name in the format 'Domain\User' or a Universal Principal Name (UPN, looks like an e-mail-address, but is not neecessarily one).

`SimulateMailboxes` is optional for simulation mode, although highly recommended. It is a comma separated list of e-mail addresses replacing the list of mailboxes otherwise gathered from the registry.

**Attention**: Simulation mode only works when the user starting the simulation is at least from the same Active Directory forest as the user defined in SimulateUser.  Users from other forests will not work.  
# 15. FAQ
## 15.1. Where can I find the changelog?
The changelog is located in the `'.\docs'` folder, along with other documents related to Set-OutlookSignatures.
## 15.2. How can I contribute, propose a new feature or file a bug?
If you have an idea for a new feature or have found a problem, please <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank">create an issue on GitHub</a>.

If you want to contribute code, please have a look at `'.\docs\CONTRIBUTING'` for a rough overview of the proposed process.
## 15.3. Why use legacyExchangeDN to find the user behind a mailbox, and not mail or proxyAddresses?  
The legacyExchangeDN attribute is used to find the user behind a mailbox, because mail and proxyAddresses are not unique in certain Exchange scenarios:  
- A separate Active Directory forest for users and Exchange mailboxes: In this case, the mail attribute is usually set in the user forest, although there are no mailboxes in this forest.  
- One common mail domain across multiple Exchange organizations: In this case, the address book is very like synchronized between Active Directory forests by using contacts or mail-enabled users, which both will have the SMTP address of the mailbox in the proxyAddresses attribute.

If Outlook is configured to access mailbox via protocols such as POP3 or IMAP4, the script searches for the legacyExchangeDN using the e-mail address of the mailbox.

Without a legacyExchangeDN, group membership information can not be retrieved. These mailboxes can still receive common and mailbox specific signatures and OOF messages.  
## 15.4. How is the personal mailbox of the currently logged in user identified?  
The personal mailbox of the currently logged in user is preferred to other mailboxes, as it receives signatures first and is the only mailbox where the Outlook Web signature can be set.

The personal mailbox is found by simply checking if the Active Directory mail attribute of the currently logged in user matches an SMTP address of one of the mailboxes connected in Outlook.

If the mail attribute is not set, the currently logged in user's objectSID is compared with all the mailboxes' msExchMasterAccountSID. If there is exactly one match, this mailbox is used as primary one.
  
Please consider the following caveats regarding the mail attribute:  
- When Active Directory attributes are directly modified to create or modify users and mailboxes (instead of using Exchange Admin Center or Exchange Management Shell), the mail attribute is often not updated and does not match the primary SMTP address of a mailbox. Microsoft strongly recommends that the mail attribute matches the primary SMTP address.  
- When using linked mailboxes, the mail attribute of the linked account is often not set or synced back from the Exchange resource forest. Technically, this is not necessary. From an organizational point of view it makes sense, as this can be used to determine if a specific user has a linked mailbox in another forest, and as some applications (such as "scan to mail") may need this attribute anyhow.  
## 15.5. Which ports are required?  
For communication with the user's own Active Directory forest, trusted domains, and their sub-domains, the following ports are usually required:
- 88 TCP/UDP (Kerberos authentication)
- 389 TCP/UPD (LDAP)
- 636 TCP (LDAP SSL)
- 3268 TCP (Global Catalog)
- 3269 TCP (Global Catalog SSL)
- 49152-65535 TCP (high ports)

The client needs the following ports to access a SMB file share on a Windows server (see <a href="https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc731402(v=ws.11)" target="_blank">this Microsoft article</a> for details):
- 137 UDP
- 138 UDP
- 139 TCP
- 445 TCP

The client needs port 443 TCP to access a WebDAV share (a SharePoint document library, for example).  
## 15.6. Why is Out of Office abbreviated OOF and not OOO?  
Back in the 1980s, Microsoft had a UNIX OS named Xenix ... but read yourself <a href="https://techcommunity.microsoft.com/t5/exchange-team-blog/why-is-oof-an-oof-and-not-an-ooo/ba-p/610191" target="_blank">here</a>.  
## 15.7. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.  
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
## 15.8. How can I log the script output?  
The script has no built-in logging option other than writing output to the host window.

You can, for example, use PowerShell's `Start-Transcript` and `Stop-Transcript` commands to create a logging wrapper around Set-OutlookSignatures.ps1.  
## 15.9. Can multiple script instances run in parallel?  
The script is designed for being run in multiple instances at the same. You can combine any of the following scenarios:  
- One user runs multiple instances of the script in parallel  
- One user runs multiple instances of the script in simulation mode in parallel  
- Multiple users on the same machine (e.g. Terminal Server) run multiple instances of the script in parallel  

Please see `'.\sample code\SimulateAndDeploy.ps1'` for an example how to run multiple instances of Set-OutlookSignatures in parallel in a controlled manner. Don't forget to adopt path names and variables to your environment.
## 15.10. How do I start the script from the command line or a scheduled task?  
Passing arguments to PowerShell.exe from the command line or task scheduler can be very tricky when spaces are involved. You have to be very careful about when to use single quotes or double quotes.

A working example:
```
PowerShell.exe -Command "& '\\server\share\directory\Set-OutlookSignatures.ps1' -SignatureTemplatePath '\\server\share\directory\templates\Signatures DOCX' -OOFTemplatePath '\\server\share\directory\templates\Out of Office DOCX' -ReplacementVariableConfigFile '\\server\share\directory\config\default replacement variables.ps1'"
```
You will find lots of information about this topic on the internet. The following links provide a first starting point:  
- <a href="https://stackoverflow.com/questions/45760457/how-can-i-run-a-powershell-script-with-white-spaces-in-the-path-from-the-command" target="_blank">https://stackoverflow.com/questions/45760457/how-can-i-run-a-powershell-script-with-white-spaces-in-the-path-from-the-command</a>
- <a href="https://stackoverflow.com/questions/28311191/how-do-i-pass-in-a-string-with-spaces-into-powershell" target="_blank">https://stackoverflow.com/questions/28311191/how-do-i-pass-in-a-string-with-spaces-into-powershell</a>
- <a href="https://stackoverflow.com/questions/10542313/powershell-and-schtask-with-task-that-has-a-space" target="_blank">https://stackoverflow.com/questions/10542313/powershell-and-schtask-with-task-that-has-a-space</a>
  
If you have to use the PowerShell.exe `-Command` or `-File` parameter depends on details of your configuration, for example AppLocker in combination with PowerShell. You may also want to consider the `-EncodedCommand` parameter to start Set-OutlookSignatures.ps1 and pass parameters to it.
  
If you provided your users a link so they can start Set-OutlookSignatures.ps1 with the correct parameters on their own, you may want to use the official icon: `'.\logo\Set-OutlookSignatures Icon.ico'`

Please see `'.\sample code\Set-OutlookSignatures.cmd'` for an example. Don't forget to adopt path names to your environment.
## 15.11. How to create a shortcut to the script with parameters?  
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

See `'.\sample code\CreateDesktopIcon.ps1'` for a code example. Don't forget to adopt path names to your environment. 
## 15.12. What is the recommended approach for implementing the software?  
There is certainly no definitive generic recommendation, but the file `'.\docs\Implementation approach.html'` should be a good starting point.

The content is based on real-life experience implementing the script in a multi-client environment with a five-digit number of mailboxes.

It contains proven procedures and recommendations for product managers, architects, operations managers, account managers and e-mail and client administrators. It is suited for service providers as well as for clients.

It covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The document is available in English and German language.  
## 15.13. What is the recommended approach for custom configuration files?
You should not change the default configuration file `'.\config\default replacement variable.ps1'`, as it might be changed in a future release of Set-OutlookSignatures. In this case, you would have to sort out the changes yourself.

The following steps are recommended:
1. Create a new custom configuration file in a separate folder.
2. The first step in the new custom configuration file should be to load the default configuration file:
   ```
   # Loading default replacement variables shipped with Set-OutlookSignatures
   . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath '\\server\share\folder\Set-OutlookSignatures\config\default replacement variables.ps1' -Raw)))
   ```
3. After importing the default configuration file, existing replacement variables can be altered with custom definitions and new replacement variables can be added.
4. Start Set-OutlookSignatures with the parameter `ReplacementVariableConfigFile` pointing to the new custom configuration file.
## 15.14. Isn't a plural noun in the script name against PowerShell best practices?
Absolutely. PowerShell best practices recommend using singular nouns, but Set-OutlookSignatures contains a plural noun.

I intentionally decided not to follow the singular noun convention, as another language as PowerShell was initially used for coding and the name of the tool was already defined. If this was a commercial enterprise project, marketing would have overruled development.
## 15.15. The script hangs at HTM/RTF export, Word shows a security warning!?
When using a signature template with account pictures (linked and embedded), conversion to HTM hangs at "Export to HTM format" or "Export to RTF format". In the background, there is a window "Microsoft Word Security Notice" with the following text:
```
Microsoft Office has identified a potential security concern.

This document contains fields that can share data with external files and websites. It is important that this file is from a trustworthy source.
```

This behavior seems new in Word versions published around August 2021. You will find several discussions regarding the message in internet forums, but I am not aware of any official statement from Microsoft.

It is yet unclear if this is a new Word security feature or a bug.

The behavior can be changed in at least two ways:
- Group Policy: Enable "User Configuration\Administrative Templates\Microsoft Word 2016\Word Options\Security\Don’t ask permission before updating IncludePicture and IncludeText fields in Word"
- Registry: Set "HKCU\SOFTWARE\Microsoft\Office\16.0\Word\Security\DisableWarningOnIncludeFieldsUpdate" (DWORD_32) to 1

Set-OutlookSignatures reads the registry key "HKCU\SOFTWARE\Microsoft\Office\16.0\Word\Security\DisableWarningOnIncludeFieldsUpdate" at start, sets it to 1 just before the conversion to HTM and RF takes place and restores the original state as soon as the conversions are finished.
This way, the warning usually gets suppressed, while the Group Policy configured state of the setting still has higher priority and overrides the user setting.
## 15.16. How to avoid empty lines when replacement variables return an empty string?
Not all users have values for all attributes, e. g. a mobile number. This can lead to empty lines in signatures, which may not look nice.

Follow these steps to avoid empty lines:
1. Use a custom replacement variable config file.
2. Modify the value of all attributes that should not leave an empty line when there is no text to show:
    - When the attribute is empty, return an empty string
    - Else, return a newline ('\`r\`n' in PowerShell) and then the attribute value.  
3. Place all required replacement variables on a single line, without a space between them.  
If they are not empty, the newline creates a new paragraph; else, the replacement variable is replaced with an emtpy string.
4. Use the ReplacementVariableConfigFile parameter when running the script.

Use '\`n' instead of '\`r\`n' to create a new line within the existing paragraph, but not a new paragraph.

When using HTML templates, use
- '\<p>' instead of '\`r\`n'
- '\<br>' instead of '\`n'

Be aware that text replacement also happens in hyperlinks ('tel:', 'mailto:' etc.).  
Instead of altering existing replacement variables, it is recommended to create new replacement variables with modified content.  
Use the new one for the pure textual replacement (including the newline), and the original one for the replacement within the hyperlink.  

The following example describes optional preceeding text combined an optional replacement variable containing a hyperlink:
- Custom replacement variable config file
  ```
  $ReplaceHash['$CURRENTUSERTELEPHONE-PREFIX-NOEMPTY$'] = $(if (-not $ReplaceHash['$CURRENTUSERTELEPHONE$']) { '' } else { "`r`nTelephone: "} )
  $ReplaceHash['$CURRENTUSERMOBILE-PREFIX-NOEMPTY$'] = $(if (-not $ReplaceHash['$CURRENTUSERMOBILE$']) { '' } else { "`r`nMobile: "} )
  ```
- Word template:  
  <a href="mailto:$CURRENTUSERMAIL$">\$CURRENTUSERMAIL\$</a>\$CURRENTUSERTELEPHONE-PREFIX-NOEMPTY\$<a href="tel:$CURRENTUSERTELEPHONE$">\$CURRENTUSERTELEPHONE\$</a>\$CURRENTUSERMOBILE-PREFIX-NOEMPTY\$<a href="tel:$CURRENTUSERMOBILE$">\$CURRENTUSERMOBILE$</a>

  Note that all variables are written on one line and that not only \$CURRENTUSERMAIL\$ is configured with a hyperlink, but \$CURRENTUSERPHONE\$ and \$CURRENTUSERMOBILE\$ too: `mailto:$CURRENTUSERMAIL$`, `tel:$CURRENTUSERTELEPHONE$` and `tel:$CURRENTUSERMOBILE$`
- Results
  - Telephone number and mobile number are set. The paragraph marks come from \$CURRENTUSERTELEPHONE-PREFIX-NOEMPTY\$ and \$CURRENTUSERMOBILE-PREFIX-NOEMPTY\$:  
    first.last@example.com  
    Telephone: <a href="tel:+43xxx">+43xxx</a>  
    Mobile: <a href="tel:+43yyy">+43yyy</a>
  - Telephone number exists, mobile number is empty. The paragraph mark comes from \$CURRENTUSERTELEPHONE-PREFIX-NOEMPTY\$:  
    first.last@example.com  
    Telephone: <a href="tel:+43xxx">+43xxx</a>
  - Telephone number is empty, mobile number is set. The paragraph mark comes from \$CURRENTUSERMOBILE-PREFIX-NOEMPTY\$  
    first.last@example.com  
    Mobile: <a href="tel:+43yyy">+43yyy</a>
## 15.17. Is there a roadmap for future versions?
There is no binding roadmap for future versions, although I maintain a list of ideas in the 'Contribution opportunities' chapter of '.\docs\CONTRIBUTING.html'.

Now that Set-OutlookSignatures is cloud aware, the next big thing will probably be supporting Microsoft's signature roaming feature. I have already seen a beta version of Outlook handling the new feature, but Microsoft has not yet disclosed an API or other detailed documentation.

Fixing issues has priority over new features, of course.
## 15.18. How to deploy signatures for "Send As", "Send On Behalf" etc.?
The script only considers primary mailboxes, these are mailboxes added as separate accounts. This is the same way Outlook handles mailboxes from a signature perspective: Outlook can not handle signatures for non-primary mailboxes (added via "Open these additional mailboxes").

If you want to deploy signatures for
- non-primary mailboxes,
- mailboxes you don't add to Outlook but just use an assigned "Send As" or "Send on Behalf" right by choosing a different "From" address,
- or distribution lists, for which you use an assigned "Send As" or "Send on Behalf" right by choosing a different "From" address,
create a group or e-mail address specific signature, where the group or the e-mail-address does not refer to the mailbox or distribution group the e-mail is sent from, but rather the user or group who has the right to send from this mailbox or distribution group.

An example:
Members of the group "Example\Group" have the right to send as mailbox m<area>@example.com and as the distribution group dg<area>@example.com.

You want to deploy signatures for the mailbox m<area>@example.com and the distribution group dg<area>@example.com.

Problem 1: dg<area>@example.com can't be added as a mailbox to Outlook, as it is a distribution group.

Problem 2: The mailbox m<area>@example.com is configured as non-primary maibox on most clients, because most of the users have the "Send as" permission, but not the "Full Access" permissions. Some users even don't connect the mailbox at all, they just choose m<area>@example.com as "From" address.

Solution: Create signature templates for the mailbox m<area>@example.com and the distribution group dg<area>@example.com and **assign them to the users and groups having "send as" permissions**.

When using file name based tags, the file names would be:
```
m@example.com external English formal.[Example Group] [u@example.com].docx

dg@example.com internal German informal.[Example Group] [u@example.com].docx
```
This works as long as the personal mailbox of a member of "Example\Group" is connected in Outlook as primary mailbox (which usually is the case). When this personal mailbox is processed by Set-OutlookSignatures, the script recognizes the group membership and the signature assigned to it.

Caveat: The \$CurrentMailbox[...]\$ replacement variables refer to the user's personal mailbox in this case, not to m<area>@example.com.
## 15.19. Can I centrally manage and deploy Outook stationery with this script?
Outlook stationery describes the layout of e-mails, including font size and color for new e-mails and for replies and forwards.

The default mail font, size and color are usually an integral part of corporate design and corporate identity. CI/CD typically also defines the content and layout of signatures.

Set-OutlookSignatures has no features regarding deploying Outlook stationery, as there are better ways for doing this.  
Outlook stores stationery settings in `'HKCU\Software\Microsoft\Office\<Version>\Common\MailSettings'`. You can use a logon script or group policies to deploy these keys, on-prem and for managed devices in the cloud.  
Unfortunately, Microsoft's group policy templates (ADMX files) for Office do not seem to provide detailed settings for Outlook stationery, so you will have to deploy registry keys. 
## 15.20. Why is membership in dynamic distribution groups and dynamic security groups not considered?
Dynamic distribution groups (DDGs) are specific groups that only work within Exchange. Group membership is evaluated just in time when an e-mail is sent to a DDG by executing the LDAP query defining a DDG.

Active Directory and Graph know that a DDG is a group, but they basically do not know the members of this group. The same is valid for dynamic security groups, which are available in the cloud only.  
In more technical words: Dynamic groups have no member attribute, and dynamic groups neither appear in the on-prem user attributes memberOf and tokenGroups nor in the Graph transitiveMemberOf query.

If dynamic groups would have to be considered, the only way would be to enumerate all dynamic groups, to run the LDAP query that defines each group, and to finally evaluate the resulting group membership.

The LDAP queries defining dynamic groups are deemed expensive due to the potential load they put on Active Directory and their resulting runtime.  
Microsoft does not recommend against dynamic groups, only not to use them heavily.  
This is very likely the reason why dynamic groups can not be granted permissions on Exchange mailboxes and other Exchange objects, and why each dynamic group can be assigned an expansion server executing the LDAP query (expansion times of 15 minutes or more are not rare in the field).

Taking all these aspects into account, Set-OutlookSignatures will not consider membership in dynamic groups until a reliable and efficient way of querying a user's dynamic group membership is available.
### 15.20.1. What's the alternative to dynamic groups?
Dynamic groups have their raison d'être, especially if you use them as a tool for special and rather rare use cases.

With the move to the cloud, where dynamic groups were introduced just not too long ago and only with a limited set of possible query parameters, an ongoing trend can be observed: Replacing dynamic groups with regularly updated static groups.

An Identity Management System (IDM) or a script regularly executes the LDAP query, which would otherwise define a dynamic group, and updates the member list of a static group.

These updates usually happen less frequent than a dynamic group is used. The static group might not be fully up-to-date when used, but other aspects outweigh this disadvantage most of the time:
- Reduced load on Active Directory (partially transferred to IDM system or server running a script)
- Static groups can be used for permissions
- Changes in static group membership can be documented more easily
- Static groups can be expanded to it's members in mail clients
- Membership in static groups can easily be queried
- Overcoming query parameter restrictions, such as combing the results of multiple LDAP queries
## 15.21. What about the new signature roaming feature Microsoft announced?  
Microsoft announced a change in how and where signatures are stored. Basically, signatures are no longer stored in the file system, but in the mailbox itself.

This is a good idea, as it makes signatures available across devices and avoids file naming conflicts which may appear in current solutions.

Based on currently available information, the disadvantage is that signatures for shared mailboxes can no longer be personalized, as the latest signature change would be propagated to all users accessing the shared mailbox (which is especially bad when personalized signatures for shared mailboxes are set as default signature).

Microsoft has stated that only cloud mailboxes support the new feature and that Outlook for Windows will be the only client supporting the new feature for now. I am confident more e-mail clients will follow soon. Future will tell if the feature will be made available for mailboxes on premises, too.

Currently, there is no detailed documentation and no API available to programatically access the new feature.

Until the feature is fully rolled out and an API is available, you can disable the feature with a registry key. This forces Outlook for Windows to use the well-known file based approach and ensures full compatibility with this script.

For details, please see <a href="https://support.microsoft.com/en-us/office/outlook-roaming-signatures-420c2995-1f57-4291-9004-8f6f97c54d15?ui=en-us&rs=en-us&ad=us" target="_blank">this Microsoft article</a>.  

### 15.21.1. Please be aware of the following problem
Since Q3 2021, the roaming signature feature appears and disappears on Outlook Web of cloud mailboxes and in  Outlook on Windows. There is still no hint of an API, or a way to disable it on the server.

When multiple signatures in Outlook Web are enabled, Set-OutlookSignatures can successfully set the signature in Outlook Web, but this signature is ignored.

There is no programmatic way to detect or change this behavior.  
The built-in Exchange Online PowerShell-Cmdlet Set-MailboxMessageConfiguration has the same problem, so it seems different Microsoft teams work on a different development and release schedule.

At the time of writing, the only known workaround is the following:
1. Delete all signatures available in Outlook Web
2. Still in Outlook Web, set the default signatures to be used for new e-mails and for replies/forwards to "(no signature)"
3. Save the updated settings
4. Wait a few minutes
5. Run Set-OutlookSignatures
6. Wait a few minutes
7. Open a new browser tab and open Outlook Web, or fully reload an existing open Outlook Web tab (Outlook Web works with caching in the browser, so it sometimes shows old configuration data) and check your signatures.

Unfortunately, further updates to the Outlook Web signature by Set-OutlookSignatures are successful but ignored by Outlook Web until all signatures are deleted manually again.

Even worse, it is not yet documented or known where the new signatures are stored and how they can be access programatically - so the deletion must happen manuelly and not be automated at the moment.

If you are affected, please let Microsoft know via a support case and https://github.com/MicrosoftDocs/office-docs-powershell/issues/8537.

As soon as there is an official API or a scriptable workaround available, Set-OutlookSignatures will be adopted to incorporate this new feature.
