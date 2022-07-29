<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="../src/logo/Set-OutlookSignatures%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-OutlookSignatures"></a>**<br>Centrally manage and deploy Outlook text signatures and Out of Office auto reply messages.<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Set-OutlookSignatures" alt=""></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Set-OutlookSignatures?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Set-OutlookSignatures/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a>  

# Features <!-- omit in toc -->
**Signatures and OOF messages can be:**
- Generated from templates in DOCX or HTML file format  
- Customized with a broad range of variables, including photos, from Active Directory and other sources  
- Applied to all mailboxes (including shared mailboxes), specific mailbox groups or specific e-mail addresses, for every primary mailbox across all Outlook profiles  
- Assigned time ranges within which they are valid  
- Set as default signature for new e-mails, or for replies and forwards (signatures only)  
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

Set-OutlookSignatures requires **no installation on servers or clients**. You only need a standard file share on a server, and PowerShell and Office. 

A **documented implementation approach**, based on real life experiences implementing the script in a multi-client environment with a five-digit number of mailboxes, contains proven procedures and recommendations for product managers, architects, operations managers, account managers and e-mail and client administrators.  
The implementatin approach is **suited for service providers as well as for clients**, and covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The script is **Free and Open-Source Software (FOSS)**. It is published under the MIT license which is approved, among others, by the Free Software Foundation (FSF) and the Open Source Initiative (OSI), and is compatible with the General Public License (GPL) v3. Please see `'.\docs\LICENSE.txt'` for copyright and MIT license details.
<br><br>
**Dear businesses using Set-OutlookSignatures:**
- Being Free and Open-Source Software, Set-OutlookSignatures can save you thousands or even tens of thousand Euros/US-Dollars per year in comparison to commercial software.  
Please consider <a href="https://github.com/sponsors/GruberMarkus" target="_blank">sponsoring this project</a> to ensure continued support, testing and enhancements.
- Invest in the open-source projects you depend on. Contributors are working behind the scenes to make open-source better for everyone - give them the help and recognition they deserve.
- Sponsor the open-source software your team has built its business on. Fund the projects that make up your software supply chain to improve its performance, reliability, and stability.
# Table of Contents <!-- omit in toc -->
- [1. Requirements](#1-requirements)
- [2. Parameters](#2-parameters)
  - [2.1. SignatureTemplatePath](#21-signaturetemplatepath)
  - [2.2. SignatureIniPath](#22-signatureinipath)
  - [2.3. ReplacementVariableConfigFile](#23-replacementvariableconfigfile)
  - [2.4. GraphConfigFile](#24-graphconfigfile)
  - [2.5. TrustsToCheckForGroups](#25-truststocheckforgroups)
  - [2.6. DeleteUserCreatedSignatures](#26-deleteusercreatedsignatures)
  - [2.7. DeleteScriptCreatedSignaturesWithoutTemplate](#27-deletescriptcreatedsignatureswithouttemplate)
  - [2.8. SetCurrentUserOutlookWebSignature](#28-setcurrentuseroutlookwebsignature)
  - [2.9. SetCurrentUserOOFMessage](#29-setcurrentuseroofmessage)
  - [2.10. OOFTemplatePath](#210-ooftemplatepath)
  - [2.11. OOFIniPath](#211-oofinipath)
  - [2.12. AdditionalSignaturePath](#212-additionalsignaturepath)
  - [2.13. UseHtmTemplates](#213-usehtmtemplates)
  - [2.14. SimulateUser](#214-simulateuser)
  - [2.15. SimulateMailboxes](#215-simulatemailboxes)
  - [2.16. GraphCredentialFile](#216-graphcredentialfile)
  - [2.17. GraphOnly](#217-graphonly)
  - [2.18. CreateRtfSignatures](#218-creatertfsignatures)
  - [2.19. CreateTxtSignatures](#219-createtxtsignatures)
  - [2.20. EmbedImagesInHtml](#220-embedimagesinhtml)
- [3. Outlook signature path](#3-outlook-signature-path)
- [4. Mailboxes](#4-mailboxes)
- [5. Group membership](#5-group-membership)
- [6. Removing old signatures](#6-removing-old-signatures)
- [7. Error handling](#7-error-handling)
- [8. Run script while Outlook is running](#8-run-script-while-outlook-is-running)
- [9. Signature and OOF file format](#9-signature-and-oof-file-format)
  - [9.1. Signature template file naming](#91-signature-template-file-naming)
- [10. Template tags and ini files](#10-template-tags-and-ini-files)
  - [10.1. Allowed tags](#101-allowed-tags)
  - [10.2. How to work with ini files](#102-how-to-work-with-ini-files)
- [11. Signature and OOF application order](#11-signature-and-oof-application-order)
- [12. Variable replacement](#12-variable-replacement)
  - [12.1. Photos from Active Directory](#121-photos-from-active-directory)
- [13. Outlook Web](#13-outlook-web)
- [14. Hybrid and cloud-only support](#14-hybrid-and-cloud-only-support)
  - [14.1. Basic Configuration](#141-basic-configuration)
  - [14.2. Advanced Configuration](#142-advanced-configuration)
  - [14.3. Authentication](#143-authentication)
- [15. Simulation mode](#15-simulation-mode)
- [16. FAQ](#16-faq)
  - [16.1. Where can I find the changelog?](#161-where-can-i-find-the-changelog)
  - [16.2. How can I contribute, propose a new feature or file a bug?](#162-how-can-i-contribute-propose-a-new-feature-or-file-a-bug)
  - [16.3. How is the account of a mailbox identified?](#163-how-is-the-account-of-a-mailbox-identified)
  - [16.4. How is the personal mailbox of the currently logged in user identified?](#164-how-is-the-personal-mailbox-of-the-currently-logged-in-user-identified)
  - [16.5. Which ports are required?](#165-which-ports-are-required)
  - [16.6. Why is Out of Office abbreviated OOF and not OOO?](#166-why-is-out-of-office-abbreviated-oof-and-not-ooo)
  - [16.7. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.](#167-should-i-use-docx-or-htm-as-file-format-for-templates-signatures-in-outlook-sometimes-look-different-than-my-templates)
  - [16.8. How can I log the script output?](#168-how-can-i-log-the-script-output)
  - [16.9. How can I get more script output for troubleshooting?](#169-how-can-i-get-more-script-output-for-troubleshooting)
  - [16.10. Can multiple script instances run in parallel?](#1610-can-multiple-script-instances-run-in-parallel)
  - [16.11. How do I start the script from the command line or a scheduled task?](#1611-how-do-i-start-the-script-from-the-command-line-or-a-scheduled-task)
  - [16.12. How to create a shortcut to the script with parameters?](#1612-how-to-create-a-shortcut-to-the-script-with-parameters)
  - [16.13. What is the recommended approach for implementing the software?](#1613-what-is-the-recommended-approach-for-implementing-the-software)
  - [16.14. What is the recommended approach for custom configuration files?](#1614-what-is-the-recommended-approach-for-custom-configuration-files)
  - [16.15. Isn't a plural noun in the script name against PowerShell best practices?](#1615-isnt-a-plural-noun-in-the-script-name-against-powershell-best-practices)
  - [16.16. The script hangs at HTM/RTF export, Word shows a security warning!?](#1616-the-script-hangs-at-htmrtf-export-word-shows-a-security-warning)
  - [16.17. How to avoid blank lines when replacement variables return an empty string?](#1617-how-to-avoid-blank-lines-when-replacement-variables-return-an-empty-string)
  - [16.18. Is there a roadmap for future versions?](#1618-is-there-a-roadmap-for-future-versions)
  - [16.19. How to deploy signatures for "Send As", "Send On Behalf" etc.?](#1619-how-to-deploy-signatures-for-send-as-send-on-behalf-etc)
  - [16.20. Can I centrally manage and deploy Outook stationery with this script?](#1620-can-i-centrally-manage-and-deploy-outook-stationery-with-this-script)
  - [16.21. Why is dynamic group membership not considered on premises?](#1621-why-is-dynamic-group-membership-not-considered-on-premises)
    - [16.21.1. Graph](#16211-graph)
    - [16.21.2. Active Directory on premises](#16212-active-directory-on-premises)
  - [16.22. Why is no admin or user GUI available?](#1622-why-is-no-admin-or-user-gui-available)
  - [16.23. What about the roaming signatures feature announced by Microsoft?](#1623-what-about-the-roaming-signatures-feature-announced-by-microsoft)
    - [16.23.1. Please be aware of the following problem](#16231-please-be-aware-of-the-following-problem)
  
# 1. Requirements  
Requires Outlook and Word, at least version 2010.  
The script must run in the security context of the currently logged in user.

The script must run in PowerShell Full Language mode. Constrained Language mode is not supported, as some features such as BASE64 conversions are not available in this mode or require very slow workarounds.

If you use AppLocker or a comparable solution, you may need to digitally sign the PowerShell 'Set-OutlokSignatures.ps1'. It is usually not necessary to sign the variable replacement configuration files, e. g. '.\config\default replacement variables.ps1'.  
There are locked down environments, where all files matching the patterns `'*.ps*1'` and `'*.dll'` need to be digitially signed with a trusted certificate. 

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
Template tags are placed in an ini file.

The file must be UTF8 encoded.

See '.\templates\Signatures DOCX\_Signatures.ini' for a sample file with further explanations.

Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures')

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'

The currently logged in user needs at least read access to the path

Default value: `'.\templates\Signatures DOCX\_Signatures.ini'`
## 2.3. ReplacementVariableConfigFile  
The parameter ReplacementVariableConfigFile tells the script where the file defining replacement variables is located.

The file must be UTF8 encoded.

Local and remote paths are supported. Local paths can be absolute (`'C:\config\default replacement variables.ps1'`) or relative to the script path (`'.\config\default replacement variables.ps1'`).

WebDAV paths are supported (https only): `'https://server.domain/SignatureSite/config/default replacement variables.ps1'` or `'\\server.domain@SSL\SignatureSite\config\default replacement variables.ps1'`

The currently logged in user needs at least read access to the file.

Default value: `'.\config\default replacement variables.ps1'`  
## 2.4. GraphConfigFile
The parameter GraphConfigFile tells the script where the file defining Graph connection and configuration options is located.

The file must be UTF8 encoded.

Local and remote paths are supported. Local paths can be absolute (`'C:\config\default graph config.ps1'`) or relative to the script path (`'.\config\default graph config.ps1'`).

WebDAV paths are supported (https only): `'https://server.domain/SignatureSite/config/default graph config.ps1'` or `'\\server.domain@SSL\SignatureSite\config\default graph config.ps1'`

The currently logged in user needs at least read access to the file.

Default value: `'.\config\default graph config.ps1'`  
## 2.5. TrustsToCheckForGroups  
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
Template tags are placed in an ini file.

The file must be UTF8 encoded.

See '.\templates\Out of Office DOCX\_OOF.ini' for a sample file with further explanations.

Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures')

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'

The currently logged in user needs at least read access to the path

Default value: `'.\templates\Out of Office DOCX\_OOF.ini'`
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
## 2.13. UseHtmTemplates  
With this parameter, the script searches for templates with the extension .htm instead of .docx.

Templates in .htm format must be UTF8 encoded.

Each format has advantages and disadvantages, please see "[13.5. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.](#135-should-i-use-docx-or-htm-as-file-format-for-templates-signatures-in-outlook-sometimes-look-different-than-my-templates)" for a quick overview.

Default value: `$false`  
## 2.14. SimulateUser  
SimulateUser is a mandatory parameter for simulation mode. This value replaces the currently logged in user.

Use a logon name in the format 'Domain\User' or a Universal Principal Name (UPN, looks like an e-mail-address, but is not neecessarily one).

See "[13. Simulation mode](#13-simulation-mode)" for details.  
## 2.15. SimulateMailboxes  
SimulateMailboxes is optional for simulation mode, although highly recommended. It is a comma separated list of e-mail addresses replacing the list of mailboxes otherwise gathered from the registry.
## 2.16. GraphCredentialFile
Path to file containing Graph credential which should be used as alternative to other token acquisition methods.

Makes only sense in combination with `'.\sample code\SimulateAndDeploy.ps1'`, do not use this parameter for other scenarios.

See `'.\sample code\SimulateAndDeploy.ps1'` for an example how to create this file.

Default value: `$null`  
## 2.17. GraphOnly
Try to connect to Microsoft Graph only, ignoring any local Active Directory.

The default behavior is to try Active Directory first and fall back to Graph.

Default value: `$false`
## 2.18. CreateRtfSignatures
Should signatures be created in RTF format?

Default value: `$true`
## 2.19. CreateTxtSignatures
Should signatures be created in TXT format?

Default value: `$true`
## 2.20. EmbedImagesInHtml
Should images be embedded into HTML files?

Outlook 2016 and newer can handle images embedded directly into an HTML file as BASE64 string (`'<img src="data:image/[...]"'`).

Outlook 2013 and earlier can't handle these embedded images when composing HTML e-mails (there is no problem receiving such e-mails, or when composing RTF or TXT e-mails).

When setting EmbedImagesInHtml to `$false`, consider setting the Outlook registry value "Send Pictures With Document" to 1 to ensure that images are sent to the recipient (see https://support.microsoft.com/en-us/topic/inline-images-may-display-as-a-red-x-in-outlook-704ae8b5-b9b6-d784-2bdf-ffd96050dfd6 for details).

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

Changing which signature is to be used as default signature for new e-mails or for replies and forwards requires restarting Outlook.   
# 9. Signature and OOF file format  
Only Word files with the extension .docx and HTML files with the extension .htm are supported as signature and OOF template files.  
## 9.1. Signature template file naming  
The name of the signature template file without extension is the name of the signature in Outlook.
Example: The template "Test signature.docx" will create a signature named "Test signature" in Outlook.

This can be overridden in the ini file with the 'OutlookSignatureName' parameter.
Example: The template "Test signature.htm" with the following ini file configuration will create a signature named "Test signature, do not use".
```
[Test signature.htm]
OutlookSignatureName = Test signature, do not use
```
# 10. Template tags and ini files
Tags define properties for templates, such as
- time ranges during which a template shall be applied or not applied
- groups whose direct or indirect members are allowed or denied application of a template
- specific e-mail addresses which are are allowed or denied application of a template
- an Outlook signature name that is different from the file name of the template
- if a signature template shall be set as default signature for new e-mails or as default signature for replies and forwards
- if a OOF template shall be set as internal or external message

There are additional tags which are not template specific, but change the behavior of Set-OutlookSignatures:
- specific sort order for templates (ascending, descending, as listed in the file)
- specific sort culture used for sorting ascendingly or descendingly (de-AT or en-US, for example)

If you want to give template creators control over the ini file, place it in the same folder as the templates.
## 10.1. Allowed tags
- `defaultNew` (signature template files only)  
    - Set signature as default signature for new mails  
- `defaultReplyFwd` (signature template files only)  
    - Set signature as default signature for replies and forwarded mails  
- `internal` (OOF template files only)  
    - Set template as default OOF message for internal recipients  
    - If neither `internal` nor `external` is defined, the template is set as default OOF message for internal and external recipients  
- `external` (OOF template files only)  
    - Set template as default OOF message for external recipients  
    - If neither `internal` nor `external` is defined, the template is set as default OOF message for internal and external recipients  
- `NetBiosDomain GroupSamAccountName`, `NetBiosDomain Display name of Group`, `-:NetBiosDomain GroupSamAccountName`, `-:NetBiosDomain Display name of Group`
  - Make this template specific for an Outlook mailbox being a direct or indirect member of this group or distribution list
  - The `'-:'` prefix makes this template invalid for the specified group.
  - Examples: `EXAMPLE Domain Users`, `-:Example GroupA`  
  - Groups must be available in Active Directory. Groups like `'Everyone'` and `'Authenticated Users'` only exist locally, not in Active Directory
  - This tag supports alternative formats, which are of special interest if you are in a cloud only or hybrid environmonent:
    - `NetBiosDomain GroupSamAccountName` and `NetBiosDomain Group DisplayName` can be queried from Microsoft Graph if the groups are synced between on-prem and the cloud. SamAccountName is queried before DisplayName. Use these formats when your environment is hybrid or on premises only.
    - `AzureAD e-mail-address-of-group@example.com`, `AzureAD GroupMailNickname`, `AzureAD GroupDisplayName` do not work with a local Active Directory, only with Microsoft Graph. They are queried in the order given. 'AzureAD' is the literal, case-insensitive string 'AzureAD', not a variable. Use these formats when you are in a cloud only environment.
  - 'NetBiosDomain' and 'EXAMPLE' are just examples. You need to replace them with the actual NetBios domain name of the Active Director domain containing the group.
  - 'AzureAD' is not an example. If you want to assign a template to a group stored in Azure Active Directory, you have to use 'AzureAD' as domain name.
  - When multiple groups are defined, membership in a single group is sufficient to be assigned the template - it is not required to be a member of all the defined groups.  
- `SmtpAddress`, `-:SmtpAddress`
  - Make this template specific for the assigned e-mail address (all SMTP addresses of a mailbox are considered, not only the primary one)
  - The `'-:'` prefix makes this template invalid for the specified e-mail address.
  - Examples: `office@example.com`, `-:test@example.com`
- `yyyyMMddHHmm-yyyyMMddHHmm`, `-:yyyyMMddHHmm-yyyyMMddHHmm`
  - Make this template valid only during the specific time range (`yyyy` = year, `MM` = month, `dd` = day, `HH` = hour (00-24), `mm` = minute).
  - The `'-:'` prefix makes this template invalid during the specified time range.
  - Examples: `202112150000-202112262359` for the 2021 Christmas season, `-:202202010000-202202282359` for a deny in February 2022
  - If the script does not run after a template has expired, the template is still available on the client and can be used.  

<br>Tags can be combined: A template may be assigned to several groups, e-mail addresses and time ranges, be denied for several groups, e-mail adresses and time ranges, be used as default signature for new e-mails and as default signature for replies and forwards - all at the same time. Simple add different tags below a file name, separated by line breaks (each tag needs to be on a separate line).

## 10.2. How to work with ini files
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
  File names not mentioned in this file are not considered, even if they are available in the file system. Set-OutlookSignatures will report files which are in the file system but not mentioned in the current ini, and vice versa.<br>  
  When there are two or more sections for a filename: The keys and values are not combined, each section is considered individually (SortCulture and SortOrder still apply).  
  This can be useful in the following scenario: Multiple shared mailboxes shall use the same template, individualized by using `$CURRENTMAILBOX[...]$` variables. A user can have multiple of these shared mailboxes in his Outlook configuration.
    - Solution A: Use multiple templates (possible in all versions)
      - Instructions
        - Create a copy of the initial template for each shared mailbox.
        - For each template copy, create a corresponding INI entry which assigns the template copy to a specific e-mail address.
      - Result
        - Templates<br>One template file for each shared mailbox
          - `template shared mailbox A.docx`
          - `template shared mailbox B.docx`
          - `template shared mailbox C.docx`
        - INI file
          ```
          [template shared mailbox A.docx]
          SharedMailboxA@example.com

          [template shared mailbox B.docx]
          SharedMailboxB@example.com

          [template shared mailbox C.docx]
          SharedMailboxC@example.com
          ```
    - Solution B: Use only one template (possible with v3.1.0 and newer)
      - Instructions
        - Create a single initial template.
        - For each shared mailbox, create a corresponding INI entry which assigns the template to a specific e-mail address and defines a separate Outlook signature name.
      - Result
        - Templates<br>One template file for all shared mailboxes
          - `template shared mailboxes.docx`
        - INI file
          ```
          [template shared mailboxes.docx]
          SharedMailboxA@example.com
          OutlookSignatureName = template SharedMailboxA

          [template shared mailboxes.docx]
          SharedMailboxB@example.com
          OutlookSignatureName = template SharedMailboxB

          [template shared mailboxes.docx]
          SharedMailboxC@example.com
          OutlookSignatureName = template SharedMailboxC
          ```
4. Add tags in the lines below the filename
  Example: `defaultNew`
    - Do not enclose tags in square brackets. This is not allowed here, but required when you add tags directly to file names.    - When an ini file is used, tags in file names are not considered as tags, but as part of the file name, so the Outlook signature name will contain them.  
    - Only one tag per line is allowed.
    Adding not a single tag to file name section is valid. The signature template is then classified as a common template.
    - Putting file names in single or double quotes is possible, but not necessary
    - Tags are case insensitive  
    `defaultNew` is the same as `DefaultNew` and `dEFAULTnEW`
    - You can override the automatic Outlook signature name generation by setting OutlookSignatureName, e. g. `OutlookSignatureName = This is a custom signature name`  
    With this option, you can have different template file names for the same Outlook signature name. Search for "Marketing external English formal" in the sample ini files for an example. Take care of signature group priorities (common, group, e-mail address) and the SortOrder and SortCulture parameters.
5. Remove the tags from the file names in the file system  
Else, the file names in the ini file and the file system do not match, which will result in some templates not being applied.  
It is recommended to create a copy of your template folder for tests.
6. Make the script use the ini file by passing the `'SignatureIniPath'` and/or `'OOFIniPath'` parameter
# 11. Signature and OOF application order  
Templates are applied in a specific order: Common tempaltes first, group templates second, e-mail address specific templates last.

Templates with a time range tag are only considered if the current system time is in range of at least one of these tags.

Common templates are templates with either no tag or only `[defaultNew]` and/or `[defaultReplyFwd]` (`[internal]` and/or `[external]` for OOF templates).

Within these groups, templates are applied alphabetically ascending.

Every centrally stored signature template is applied only once, as there is only one signature path in Outlook, and subfolders are not allowed - so the file names have to be unique.

The script always starts with the mailboxes in the default Outlook profile, preferrably with the current users personal mailbox.

OOF templates are only applied if the Out of Office assistant is currently disabled. If it is currently active or scheduled to be activated in the future, OOF templates are not applied.  
# 12. Variable replacement  
Variables are case sensitive.

Variables are replaced everywhere, including links, QuickTips and alternative text of images.

With this feature, you can not only show e-mail addresses and telephone numbers in the signature and OOF message, but show them as links which open a new e-mail message (`"mailto:"`) or dial the number (`"tel:"`) via a locally installed softphone when clicked.

Custom Active directory attributes are supported as well as custom replacement variables, see `'.\config\default replacement variables.ps1'` for details.  
Attributes from Microsoft Graph need to be mapped, this is done in `'.\config\default graph config.ps1'`.

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
    - `$CURRENTUSERSTATE$`: State  
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
## 12.1. Photos from Active Directory  
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
# 13. Outlook Web  
If the currently logged in user has configured his personal mailbox in Outlook, the default signature for new emails is configured in Outlook Web automatically.

If the default signature for new mails matches the one used for replies and forwarded e-mail, this is also set in Outlook.

If different signatures for new and reply/forward are set, only the new signature is copied to Outlook Web.

If only a default signature for replies and forwards is set, only this new signature is copied to Outlook Web.

If there is no default signature in Outlook, Outlook Web settings are not changed.

All this happens with the credentials of the currently logged in user, without any interaction neccessary.  
# 14. Hybrid and cloud-only support
Set-OutlookSignatures supports three directory environments:
- Active Directory on premises. This requires direct connection to Active Directory Domain Controllers, which usually only works when you are connected to your company network.
- Hybrid. This environment consists of an Active Directory on premises, which is synced with Microsoft 365 Azure Active Directory in the cloud. If the script can't make a connection to your on-prem environment, it tries to get required data from the cloud via the Microsoft Graph API.
- Cloud-only. This environment has no Active Directory on premises, only Microsoft 365 with Azure Active Directory is used. If the script can't make a connection to your on-prem environment, it tries to get required data from the cloud via the Microsoft Graph API.
## 14.1. Basic Configuration
To allow communication between Microsoft Graph and Set-Outlooksignatures, both need to be configured for each other.

The easiest way is to once start Set-OutlookSignatures with a cloud administrator. The administrator then gets asked for admin consent for the correct permissions.  
If you don't want to use custom Graph attributes or other advanced configurations, no more configuration in Microsoft Graph or Set-OutlookSignatures is required.

If you prefer using own application IDs or need advanced configuration, follow these steps:  
- In Microsoft Graph, with an administrative account:
  - Create an application with a Client ID
  - Provide admin consent (pre-approval) for the following scopes (permissions):
    - '`https://graph.microsoft.com/openid`' for logging-on the user
    - '`https://graph.microsoft.com/email`' for reading the logged-on user's mailbox properties
    - '`https://graph.microsoft.com/profile`' for reading the logged-on user's properties
    - '`https://graph.microsoft.com/user.read.all`' for reading properties of other users (manager, additional mailboxes and their managers)
    - '`https://graph.microsoft.com/group.read.all`' for reading properties of all groups, required for templates restricted to groups
    - '`https://graph.microsoft.com/mailboxsettings.readwrite`' for updating the logged-on user's Out of Office auto reply messages
    - '`https://graph.microsoft.com/EWS.AccessAsUser.All`' for updating the logged-on user's Outlook Web signature
  - Set the Redirect URI to '`https://localhost`' and configure it for '`mobile and desktop applications`'
  - Enable '`Allow public client flows`' to make Windows Integrated Authentication (SSO) work for Azure AD joined devices
- In Set-OutlookSignature, use '`.\config\default graph config.ps1`' as a template for a custom Graph configuration file
  - Set '`$GraphClientID`' to the application ID created by the Graph administrator before
  - Use the '`GraphConfigFile`' parameter to make the tool use the newly created Graph configuration file.
## 14.2. Advanced Configuration
The Graph configuration file allows for additional, advanced configuration:
- `$GraphEndpointVersion`: The version of the Graph REST API to use
- `$GraphUserProperties`: The properties to load for each graph user/mailbox. You can add custom attributes here.
- `$GraphUserAttributeMapping`: Graph and Active Directory attributes are not named identically. Set-OutlookSignatures therefore uses a "virtual" account. Use this hashtable to define which Graph attribute name is assigned to which attribute of the virtual account.  
The virtual account is accessible as `$ADPropsCurrentUser[...]` in `'.\config\default replacement variables.ps1'`, and therefore has a direct impact on replacement variables.
## 14.3. Authentication
In hybrid and cloud-only scenarios, Set-OutlookSignatures automatically tries three stages of authentication.
1. Windows Integrated Authentication  
  This works in hybrid scenarios. The credentials of the currently logged in user are used to access Microsoft Graph without any further user interaction.
2. Silent authentication  
  If Windows Integrated Authentication fails, the User Principal Name of the currently logged in user is determined. If an existing cached cloud credential for this UPN is found, it is used for authentication with Microsoft Graph.  
  A default browser window with an "Authentication successful" message may open, it can be closed anytime.
3. User interaction  
  If the other authentication methods fail, the user is interactively asked for credentials. No custom components are used, only the official Microsoft 365 authentication site and the user's default browser. 
# 15. Simulation mode  
Simulation mode is enabled when the parameter SimulatedUser is passed to the script. It answers the question `"What will the signatures look like for user A, when Outlook is configured for the mailboxes X, Y and Z?"`.

Simulation mode is useful for content creators and admins, as it allows to simulate the behavior of the script and to inspect the resulting signature files before going live.
  
In simulation mode, Outlook registry entries are not considered and nothing is changed in Outlook and Outlook web.

The template files are handled just as during a real script run, but only saved to the folder passed by the parameters AdditionalSignaturePath and AdditionalSignaturePath folder.
  
`SimulateUser` is a mandatory parameter for simulation mode. This value replaces the currently logged in user. Use a logon name in the format 'Domain\User' or a Universal Principal Name (UPN, looks like an e-mail address, but is not neecessarily one).

`SimulateMailboxes` is optional for simulation mode, although highly recommended. It is a comma separated list of e-mail addresses replacing the list of mailboxes otherwise gathered from the registry.

**Attention**: Simulation mode only works when the user starting the simulation is at least from the same Active Directory forest as the user defined in SimulateUser.  Users from other forests will not work.  
# 16. FAQ
## 16.1. Where can I find the changelog?
The changelog is located in the `'.\docs'` folder, along with other documents related to Set-OutlookSignatures.
## 16.2. How can I contribute, propose a new feature or file a bug?
If you have an idea for a new feature or have found a problem, please <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank">create an issue on GitHub</a>.

If you want to contribute code, please have a look at `'.\docs\CONTRIBUTING'` for a rough overview of the proposed process.
## 16.3. How is the account of a mailbox identified?
The legacyExchangeDN attribute is the preferred method to find the account of a mailbox, as this also works in specific scenarios where the mail and proxyAddresses attribute is not sufficient:
- Separate Active Directory forests for users and Exchange mailboxes: In this case, the mail attribute is usually set in the user forest, although there are no mailboxes in this forest.
- One common e-mail domain across multiple Exchange organizations: In this case, the address book is very like synchronized between Active Directory forests by using contacts or mail-enabled users, which both will have the SMTP address of the mailbox in the proxyAddresses attribute.

The legacyExchangeDN search considers migration scenarios where the original legacyExchangeDN is only available as X500 address in the proxyAddresses attribute of the migrated mailbox, or where the the mailbox in the source system has been converted to a mail enabled user still having the old legacyExchangeDN attribute.

If Outlook does not have information about the legacyExchangeDN of a mailbox (for example, when accessing a mailbox via protocols such as POP3 or IMAP4), the account behind a mailbox is searched by checking if the e-mail address of the mailbox can be found in the proxyAddresses attribute of an account in Active Directory/Graph.

If the account behind a mailbox is found, group membership information be retrieved and group specific templates can be applied.
If the account behind a mailbox is not found, group membership can not be retrieved and group specific templates can not be applied. Such mailboxes can still receive common and mailbox specific signatures and OOF messages.  
## 16.4. How is the personal mailbox of the currently logged in user identified?  
The personal mailbox of the currently logged in user is preferred to other mailboxes, as it receives signatures first and is the only mailbox where the Outlook Web signature can be set.

The personal mailbox is found by simply checking if the Active Directory mail attribute of the currently logged in user matches an SMTP address of one of the mailboxes connected in Outlook.

If the mail attribute is not set, the currently logged in user's objectSID is compared with all the mailboxes' msExchMasterAccountSID. If there is exactly one match, this mailbox is used as primary one.
  
Please consider the following caveats regarding the mail attribute:  
- When Active Directory attributes are directly modified to create or modify users and mailboxes (instead of using Exchange Admin Center or Exchange Management Shell), the mail attribute is often not updated and does not match the primary SMTP address of a mailbox. Microsoft strongly recommends that the mail attribute matches the primary SMTP address.  
- When using linked mailboxes, the mail attribute of the linked account is often not set or synced back from the Exchange resource forest. Technically, this is not necessary. From an organizational point of view it makes sense, as this can be used to determine if a specific user has a linked mailbox in another forest, and as some applications (such as "scan to e-mail") may need this attribute anyhow.  
## 16.5. Which ports are required?  
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
## 16.6. Why is Out of Office abbreviated OOF and not OOO?  
Back in the 1980s, Microsoft had a UNIX OS named Xenix ... but read yourself <a href="https://techcommunity.microsoft.com/t5/exchange-team-blog/why-is-oof-an-oof-and-not-an-ooo/ba-p/610191" target="_blank">here</a>.  
## 16.7. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.  
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
Get-ChildItem '.\templates\Signatures HTML' -File | ForEach-Object {
    $_.FullName  
    ConvertTo-SingleFileHTML $_.FullName ($_.FullName -replace '.htm$', ' embedded.htm')
}
```

The templates delivered with this script represent all possible formats:  
- `'.\templates\Out of Office DOCX'` and `'.\templates\signatures DOCX'` contain templates in the DOCX format  
- `'.\templates\Out of Office HTML'` contains templates in the HTML format as Word exports them when using `"Website, filtered"` as format. Note the additional folders for each signature.  
- `'.\templates\Signatures HTML'` contains templates in the HTML format. Note that there are no additional folders, as the Word export files have been processed with ConvertTo-SingleFileHTML function to create a single HTMl file with all local images embedded.  
## 16.8. How can I log the script output?  
The script has no built-in logging option other than writing output to the host window.

You can, for example, use PowerShell's `Start-Transcript` and `Stop-Transcript` commands to create a logging wrapper around Set-OutlookSignatures.ps1.  
## 16.9. How can I get more script output for troubleshooting?
Start the script with the '-verbose' parameter to get the maximum output for troubleshooting.
## 16.10. Can multiple script instances run in parallel?  
The script is designed for being run in multiple instances at the same. You can combine any of the following scenarios:  
- One user runs multiple instances of the script in parallel  
- One user runs multiple instances of the script in simulation mode in parallel  
- Multiple users on the same machine (e.g. Terminal Server) run multiple instances of the script in parallel  

Please see `'.\sample code\SimulateAndDeploy.ps1'` for an example how to run multiple instances of Set-OutlookSignatures in parallel in a controlled manner. Don't forget to adopt path names and variables to your environment.
## 16.11. How do I start the script from the command line or a scheduled task?  
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
## 16.12. How to create a shortcut to the script with parameters?  
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
## 16.13. What is the recommended approach for implementing the software?  
There is certainly no definitive generic recommendation, but the file `'.\docs\Implementation approach.html'` should be a good starting point.

The content is based on real life experiences implementing the script in a multi-client environment with a five-digit number of mailboxes.

It contains proven procedures and recommendations for product managers, architects, operations managers, account managers and e-mail and client administrators. It is suited for service providers as well as for clients.

It covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The document is available in English and German language.  
## 16.14. What is the recommended approach for custom configuration files?
You should not change the default configuration files `'.\config\default replacement variable.ps1'` and `'.\config\default graph config.ps1'`, as they might be changed in a future release of Set-OutlookSignatures. In this case, you would have to sort out the changes yourself.

The following steps are recommended:
1. Create a new custom configuration file in a separate folder.
2. The first step in the new custom configuration file should be to load the default configuration file, `'.\config\default replacement variable.ps1'` in this example:
   ```
   # Loading default replacement variables shipped with Set-OutlookSignatures
   . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath '\\server\share\folder\Set-OutlookSignatures\config\default replacement variables.ps1' -Raw)))
   ```
3. After importing the default configuration file, existing replacement variables can be altered with custom definitions and new replacement variables can be added.
4. Start Set-OutlookSignatures with the parameter `ReplacementVariableConfigFile` pointing to the new custom configuration file.
## 16.15. Isn't a plural noun in the script name against PowerShell best practices?
Absolutely. PowerShell best practices recommend using singular nouns, but Set-OutlookSignatures contains a plural noun.

I intentionally decided not to follow the singular noun convention, as another language as PowerShell was initially used for coding and the name of the tool was already defined. If this was a commercial enterprise project, marketing would have overruled development.
## 16.16. The script hangs at HTM/RTF export, Word shows a security warning!?
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
## 16.17. How to avoid blank lines when replacement variables return an empty string?
Not all users have values for all attributes, e. g. a mobile number. These empty attributes can lead to blank lines in signatures, which may not look nice.

Follow these steps to avoid blank lines:
1. Use a custom replacement variable config file.
2. Modify the value of all attributes that should not leave an blank line when there is no text to show:
    - When the attribute is empty, return an empty string
    - Else, return a newline (`Shift+Enter` in Word, `` `n `` in PowerShell, `<br>` in HTML) or a paragraph mark (`Enter` in Word, `` `r`n `` in PowerShell, `<p>` in HTML), and then the attribute value.  
3. Place all required replacement variables on a single line, without a space between them. The replacement variables themselves contain the requires newline or paragraph marks.
4. Use the ReplacementVariableConfigFile parameter when running the script.

Be aware that text replacement also happens in hyperlinks (`tel:`, `mailto:` etc.).  
Instead of altering existing replacement variables, it is recommended to create new replacement variables with modified content.  
Use the new one for the pure textual replacement (including the newline), and the original one for the replacement within the hyperlink.  

The following example describes optional preceeding text combined with an optional replacement variable containing a hyperlink.  
The internal variable `$UseHtmTemplates` is used to automatically differentiate between DOCX and HTM line breaks.
- Custom replacement variable config file
  ```
  $ReplaceHash['$CURRENTUSERTELEPHONE-PREFIX-NOEMPTY$'] = $(if (-not $ReplaceHash['$CURRENTUSERTELEPHONE$']) { '' } else { $(if ($UseHtmTemplates) { '<br>' } else { "`n" }) + 'Telephone: ' } )
  $ReplaceHash['$CURRENTUSERMOBILE-PREFIX-NOEMPTY$'] = $(if (-not $ReplaceHash['$CURRENTUSERMOBILE$']) { '' } else { $(if ($UseHtmTemplates) { '<br>' } else { "`n" }) + 'Mobile: ' } )
  ```
- Word template:  
  <pre><code>E-Mail: <a href="mailto:$CURRENTUSERMAIL$">$CURRENTUSERMAIL$</a>$CURRENTUSERTELEPHONE-PREFIX-NOEMPTY$<a href="tel:$CURRENTUSERTELEPHONE$">$CURRENTUSERTELEPHONE$</a>$CURRENTUSERMOBILE-PREFIX-NOEMPTY$<a href="tel:$CURRENTUSERMOBILE$">$CURRENTUSERMOBILE$</a></code></pre>

  Note that all variables are written on one line and that not only `$CURRENTUSERMAIL$` is configured with a hyperlink, but `$CURRENTUSERPHONE$` and `$CURRENTUSERMOBILE$` too:
  - `mailto:$CURRENTUSERMAIL$`
  - `tel:$CURRENTUSERTELEPHONE$`
  - `tel:$CURRENTUSERMOBILE$`
- Results
  - Telephone number and mobile number are set.  
  The paragraph marks come from `$CURRENTUSERTELEPHONE-PREFIX-NOEMPTY$` and `$CURRENTUSERMOBILE-PREFIX-NOEMPTY$`.  
    <pre><code>E-Mail: <a href="mailto:first.last@example.com">first.last@example.com</a>
    Telephone: <a href="tel:+43xxx">+43xxx</a>
    Mobile: <a href="tel:+43yyy">+43yyy</a></code></pre>
  - Telephone number is set, mobile number is empty.  
  The paragraph mark comes from `$CURRENTUSERTELEPHONE-PREFIX-NOEMPTY$`.  
    <pre><code>E-Mail: <a href="mailto:first.last@example.com">first.last@example.com</a>
    Telephone: <a href="tel:+43xxx">+43xxx</a></code></pre>
  - Telephone number is empty, mobile number is set.  
  The paragraph mark comes from `$CURRENTUSERMOBILE-PREFIX-NOEMPTY$`.  
    <pre><code>E-Mail: <a href="mailto:first.last@example.com">first.last@example.com</a>
    Mobile: <a href="tel:+43yyy">+43yyy</a></code></pre>
## 16.18. Is there a roadmap for future versions?
There is no binding roadmap for future versions, although I maintain a list of ideas in the 'Contribution opportunities' chapter of '.\docs\CONTRIBUTING.html'.

Fixing issues has priority over new features, of course.
## 16.19. How to deploy signatures for "Send As", "Send On Behalf" etc.?
The script only considers primary mailboxes, these are mailboxes added as separate accounts. This is the same way Outlook handles mailboxes from a signature perspective: Outlook can not handle signatures for non-primary mailboxes (added via "Open these additional mailboxes").

If you want to deploy signatures for
- non-primary mailboxes,
- mailboxes you don't add to Outlook but just use an assigned "Send As" or "Send on Behalf" right by choosing a different "From" address,
- or distribution lists, for which you use an assigned "Send As" or "Send on Behalf" right by choosing a different "From" address,
create a group or e-mail address specific signature, where the group or the e-mail address does not refer to the mailbox or distribution group the e-mail is sent from, but rather the user or group who has the right to send from this mailbox or distribution group.

An example:
Members of the group "Example\Group" have the right to send as mailbox m<area>@example.com and as the distribution group dg<area>@example.com.

You want to deploy signatures for the mailbox m<area>@example.com and the distribution group dg<area>@example.com.

Problem 1: dg<area>@example.com can't be added as a mailbox to Outlook, as it is a distribution group.

Problem 2: The mailbox m<area>@example.com is configured as non-primary maibox on most clients, because most of the users have the "Send as" permission, but not the "Full Access" permissions. Some users even don't connect the mailbox at all, they just choose m<area>@example.com as "From" address.

Solution: Create signature templates for the mailbox m<area>@example.com and the distribution group dg<area>@example.com and **assign them to the group that has been granted the "send as" permission**:
```
[External English formal m@example.com.docx]
Example Group

[External English formal dg@example.com.docx]
Example Group
```
This works as long as the personal mailbox of a member of "Example\Group" is connected in Outlook as primary mailbox (which usually is the case). When this personal mailbox is processed by Set-OutlookSignatures, the script recognizes the group membership and the signature assigned to it.

Caveat: The `$CurrentMailbox[...]$` replacement variables refer to the user's personal mailbox in this case, not to m<area>@example.com.
## 16.20. Can I centrally manage and deploy Outook stationery with this script?
Outlook stationery describes the layout of e-mails, including font size and color for new e-mails and for replies and forwards.

The default e-mail font, size and color are usually an integral part of corporate design and corporate identity. CI/CD typically also defines the content and layout of signatures.

Set-OutlookSignatures has no features regarding deploying Outlook stationery, as there are better ways for doing this.  
Outlook stores stationery settings in `'HKCU\Software\Microsoft\Office\<Version>\Common\MailSettings'`. You can use a logon script or group policies to deploy these keys, on-prem and for managed devices in the cloud.  
Unfortunately, Microsoft's group policy templates (ADMX files) for Office do not seem to provide detailed settings for Outlook stationery, so you will have to deploy registry keys. 
## 16.21. Why is dynamic group membership not considered on premises?
Membership in dynamic groups, no matter if they are of the security or distribution type, is considered only when using Microsoft Graph.

Dynamic group membership is not considered when using an on premises Active Directory. 

The reason for this is that Graph and on-prem AD handle dynamic group membership differently:
### 16.21.1. Graph
Microsoft Graph caches information about dynamic group membership at the group as well as at the user level.  

Graph regularly executes the LDAP queries defining dynamic groups and updates existing attributes with member information.  
Dynamic groups in Graph are therefore not strictly dynamic in terms of running the defining LDAP query every time a dynamic group is used and thus providing near real-time member information - they behave more like regularly updated static groups, which makes handling for scripts and applications much easier.

For the usecases of Set-OutlookSignatures, there is no difference between a static and a dynamic group in Graph:
- Querying the '`transitiveMemberOf`' attribute of a user returns static as well as dynamic group membership.
- Querying the '`members`' attribute of a group returns the group's members, no matter if the group is static or dynamic.
### 16.21.2. Active Directory on premises
Active Directory on premises does not cache any information about membership in dynamic groups at the user level, so dynamic groups do not appear in attributes such as '`memberOf`' and '`tokenGroups`'.  
Active Directory on premises also does not cache any information about members of dynamic groups at the group level, so the group attribute '`members`' is always empty. 

If dynamic groups would have to be considered, the only way would be to enumerate all dynamic groups, to run the LDAP query that defines each group, and to finally evaluate the resulting group membership.

The LDAP queries defining dynamic groups are deemed expensive due to the potential load they put on Active Directory and their resulting runtime.  
Microsoft does not recommend against dynamic groups, only not to use them heavily.  
This is very likely the reason why dynamic groups can not be granted permissions on Exchange mailboxes and other Exchange objects, and why each dynamic group can be assigned an expansion server executing the LDAP query (expansion times of 15 minutes or more are not rare in the field).

Taking all these aspects into account, Set-OutlookSignatures will not consider membership in dynamic groups on premises until a reliable and efficient way of querying a user's dynamic group membership is available.

A possible way around this restriction is replacing dynamic groups with regularly updated static groups (which is what Microsoft Graph does automatically in the background):
- An Identity Management System (IDM) or a script regularly executes the LDAP query, which would otherwise define a dynamic group, and updates the member list of a static group.
- These updates usually happen less frequent than a dynamic group is used. The static group might not be fully up-to-date when used, but other aspects outweigh this disadvantage most of the time:
  - Reduced load on Active Directory (partially transferred to IDM system or server running a script)
  - Static groups can be used for permissions
  - Changes in static group membership can be documented more easily
  - Static groups can be expanded to it's members in e-mail clients
  - Membership in static groups can easily be queried
  - Overcoming query parameter restrictions, such as combining the results of multiple LDAP queries
## 16.22. Why is no admin or user GUI available?
From an admin perspective, Set-OutlookSignatures has been designed to work with on-board tools wherever possible and to make managing and deploying signatures intuitive.

This "easy to set up, easy to understand, easy to maintain" approach is why
- there is no need for a dedicated server, a database or a setup program
- Word documents are supported as templates in addition to HTML templates
- there is the clear hierarchy of common, group specific and e-mail address specific template application order

For an admin, the most complicated part is bringing Set-OutlookSignatures to his users by integrating it into the logon script, deploy a desktop icon or start menu entry, or creating a scheduled task. Alternatively, an admin can use a signature deployment method without user or client involvement.  
Both tasks are usually neccessary only once, sample code and documentation based on real life experiences are available.  
Anyhow, a basic GUI for configuring the script is accessible via the following built-in PowerShell command:
```
Show-Command .\Set-OutlookSignatures.ps1
```

For a template creator/maintainer, maintaining the INI files defining template application order and permissions is the main task, in combination with tests using simulation mode.  
These tasks typically happen multiple times a year. A graphical user interface might make them more intuitive and easier; until then, documentation and examples based on real life experiences are available.

From an end user perspective, Set-OutlookSignatures should not have a GUI at all. It should run in the background or on demand, but there should be no need for any user interaction.

## 16.23. What about the roaming signatures feature announced by Microsoft?  
Microsoft announced a future change in how and where signatures are stored. Basically, signatures will no longer stored in the file system, but in the mailbox itself.  
For details, please see <a href="https://support.microsoft.com/en-us/office/outlook-roaming-signatures-420c2995-1f57-4291-9004-8f6f97c54d15?ui=en-us&rs=en-us&ad=us" target="_blank">this Microsoft article</a>.  

This is a good idea, as it makes signatures available across devices and apps.

Some personal educated guesses based on available documentation, Outlook for Windows beta versions and several Exchange Online tenants:
- The feature has first been annount by Microsoft in 2020, but has been postponed multiple times. At the time of writing this, the feature shall be released publicly in October 2022 according to the Office 365 roadmap.
- Microsoft has not yet published a public API. 
- Outlook for Windows is the only client mentioned to support the new feature for now. I am confident more e-mail clients - especially Outlook for Mac, iOS and Android - will follow (the sooner, the better).
- The roaming signatures feature will very likely only be available for mailboxes in the cloud. Mailboxes on on-prem servers will not support this feature, no matter if in pure on-prem or in hybrid scenarios.
- It is yet unclear if this feature will be available for shared mailboxes. If yes, the disadvantage is that signatures for shared mailboxes can no longer be personalized, as the latest signature change would be propagated to all users accessing the shared mailbox (which is especially bad when personalized signatures for shared mailboxes are set as default signature - think about '`$CURRENTUSER[...]$`' replacement variables).

Outlook for Windows beta versions already support the roaming signatures feature. Until the feature is fully rolled out and an API is available, you can disable the feature with a registry key. This forces Outlook for Windows to use the well-known file based approach and ensure full compatibility with Set-OutlookSignatures, until a public API is released and incorporated into the script.
  - With the '`DisableRoamingSignaturesTemporaryToggle`' registry value being absent or set to 0, file based signatures created by tools such as Set-OutlookSignatures are regularly deleted and replaced with signatures stored directly in the mailbox.
  - With the '`DisableRoamingSignaturesTemporaryToggle`' registry value set to 1, the file based approach continues to work as known. Outlook does not synchronize signatures to the mailbox.

Microsoft is already supporting the feature in Outlook Web for more and more Exchange Online tenants. Currently, this breaks PowerShell commands such as Set-MailboxMessageConfiguration and there is no public API available.
  - Set-OutlookSignatures can set one Outlook Web signature, but an Exchange Online tenant with multiple signatures feature enabled just ignores this signature (see the next chapter for workarounds).
### 16.23.1. Please be aware of the following problem
Since Q3 2021, the roaming signature feature appears and disappears on Outlook Web of cloud mailboxes. There is still no hint of an API, or a way to disable it on the server.

When multiple signatures in Outlook Web are enabled, Set-OutlookSignatures can successfully set the signature in Outlook Web, but this signature is ignored.

There is no programmatic way to detect or change this behavior.  
The built-in Exchange Online PowerShell-Cmdlet '`Set-MailboxMessageConfiguration`' has the same problem, so it seems different Microsoft teams work on a different development and release schedule.

At the time of writing, there are two workarounds:
- Manual approach
  1. Delete all signatures available in Outlook Web
  2. Still in Outlook Web, set the default signatures to be used for new e-mails and for replies/forwards to "(no signature)"
  3. Save the updated settings
  4. Wait a few minutes
  5. Run Set-OutlookSignatures
  6. Wait a few minutes
  7. Open a new browser tab and open Outlook Web, or fully reload an existing open Outlook Web tab (Outlook Web works with caching in the browser, so it sometimes shows old configuration data) and check your signatures.
  8. Unfortunately, further updates to the Outlook Web signature by Set-OutlookSignatures are successful but ignored by Outlook Web until all signatures are deleted manually again. Even worse, it is not yet documented or known where the new signatures are stored and how they can be access programatically - so the deletion must happen manuelly and can not be automated at the moment.
- Disable the feature in your tenant
  - Only Microsoft can do this. Let Microsoft know via a support case.

As soon as there is an official API or a scriptable workaround available, it will be evaluated for support in Set-OutlookSignatures.
