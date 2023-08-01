<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="../src_Set-OutlookSignatures/logo/Set-OutlookSignatures%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-OutlookSignatures"></a>**<br>Centrally manage and deploy Outlook text signatures and Out of Office auto reply messages<p><p><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt="this release"></a> XXXRemoveWhenBuildingXXX--> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Set-OutlookSignatures?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="latest release" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Set-OutlookSignatures" alt="open issues" data-external="1"></a> <img src="https://img.shields.io/badge/dynamic/json?color=brightgreen&label=views&query=%24.count&url=https%3A%2F%2Fraw.githubusercontent.com%2FGruberMarkus%2Fmy-traffic2badge%2Ftraffic%2Ftraffic-Set-OutlookSignatures%2Ftraffic_views.json" alt="views" data-external="1"> <img src="https://img.shields.io/badge/dynamic/json?color=brightgreen&label=clones&query=%24.count&url=https%3A%2F%2Fraw.githubusercontent.com%2FGruberMarkus%2Fmy-traffic2badge%2Ftraffic%2Ftraffic-Set-OutlookSignatures%2Ftraffic_clones.json" alt="clones" data-external="1"> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Set-OutlookSignatures?color=brightgreen" alt="stars" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/donate,%20support,%20sponsor-white?logo=githubsponsors" alt="donate or sponsor"></a> <a href="./Benefactor%20Circle.md" target="_blank"><img src="https://img.shields.io/badge/unlock%20all%20features%20with-Benefactor%20Circle-gold" alt="unlock all features with Benefactor Circle"></a>
# Features <!-- omit in toc -->
**Signatures and OOF messages can be:**
- Generated from **templates in DOCX or HTML** file format  
- Customized with a **broad range of variables**, including **photos**, from Active Directory and other sources
  - Variables are available for the **currently logged-on user, this user's manager, each mailbox and each mailbox's manager**
  - Images in signatures can be **bound to the existence of certain variables** (useful for optional social network icons, for example)
- Applied to all **mailboxes (including shared mailboxes)**, specific **mailbox groups**, specific **e-mail addresses** or specific **user or mailbox properties**, for **every mailbox across all Outlook profiles** (**automapped and additional mailboxes** are optional)  
- Created with different names from the same template (e.g., **one template can be used for multiple shared mailboxes**)
- Assigned **time ranges** within which they are valid  
- Set as **default signature** for new e-mails, or for replies and forwards (signatures only)  
- Set as **default OOF message** for internal or external recipients (OOF messages only)  
- Set in **Outlook Web** for the currently logged-in user  
- Centrally managed only or **exist along user created signatures** (signatures only)  
- Copied to an **alternate path** for easy access on mobile devices not directly supported by this script (signatures only)
- **Write protected** (Outlook signatures only)
- Mirrored to the cloud as **roaming signatures**

Set-Outlooksignatures can be **executed by users on clients, or on a server without end user interaction**.  
On clients, it can run as part of the logon script, as scheduled task, or on user demand via a desktop icon, start menu entry, link or any other way of starting a program.  
Signatures and OOF messages can also be created and deployed centrally, without end user or client involvement.

**Sample templates** for signatures and OOF messages demonstrate all available features and are provided as .docx and .htm files.

**Simulation mode** allows content creators and admins to simulate the behavior of the script and to inspect the resulting signature files before going live.
  
The script is **designed to work in big and complex environments** (Exchange resource forest scenarios, across AD trusts, multi-level AD subdomains, many objects). It works **on premises, in hybrid and cloud-only environments**.

It is **multi-client capable** by using different template paths, configuration files and script parameters.

Set-OutlookSignatures requires **no installation on servers or clients**. You only need a standard file share on a server, and PowerShell and Office. 

A **documented implementation approach**, based on real life experiences implementing the script in multi-client environments with a five-digit number of mailboxes, contains proven procedures and recommendations for product managers, architects, operations managers, account managers and e-mail and client administrators.  
The implementatin approach is **suited for service providers as well as for clients**, and covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The script core is **Free and Open-Source Software (FOSS)**. It is published under the MIT license which is approved, among others, by the Free Software Foundation (FSF) and the Open Source Initiative (OSI), and is compatible with the General Public License (GPL) v3. Please see `.\docs\LICENSE.txt` for copyright and MIT license details.

**Some features are exclusive to Benefactor Circle members.** Benefactor Circle members have access to an extension file enabling the exclusive features. This extension file is chargeable, and it is distributed under a proprietary, non-free and non-open-source licence.  Please see `.\docs\Benefactor Circle` for details.  

**A big "Thank you!" for listing, featuring, supporting or sponsoring Set-OutlookSignatures!**
<pre><a href="https://explicitconsulting.at" target="_blank"><img src="../src_Set-OutlookSignatures/logo/Others/ExplicIT Consulting, color on black.png" height="100" title="ExplicIT Consulting" alt="ExplicIT Consulting"></a>  <a href="https://joinup.ec.europa.eu/collection/free-and-open-source-software/solution/set-outlooksignatures/about" target="_blank"><img src="../src_Set-OutlookSignatures/logo/Others/EC Joinup Interoperable Europe.png" height="100" title="European Commission Joinup/Interoperable Europe programs" alt="European Commission Joinup/Interoperable Europe programs"></a>  <a href="https://startups.microsoft.com" target="_blank"><img src="../src_Set-OutlookSignatures/logo/Others/MS_Startups_Celebration_Badge_Dark.png" height="100" title="Proud to partner with Microsoft for Startups" alt="Proud to partner with Microsoft for Startups"></a>  <a href="https://archiveprogram.github.com/" target="_blank"><img src="../src_Set-OutlookSignatures/logo/Others/GitHub-Archive-Program-logo.png" height="100" title="GitHub Archive Program" alt="GitHub Archive Program"></a></pre>

## Demo video  <!-- omit in toc -->
<a href="https://www.youtube-nocookie.com/embed/K9TrCjTdRUI" target="_blank"><img src="https://img.youtube.com/vi/K9TrCjTdRUI/hqdefault.jpg" height="300" title="Set-OutlookSignatures demo video" alt="Set-OutlookSignatures demo video"></a>

# Table of Contents <!-- omit in toc -->
- [1. Requirements](#1-requirements)
- [2. Quick Start Guide](#2-quick-start-guide)
- [3. Parameters](#3-parameters)
  - [3.1. SignatureTemplatePath](#31-signaturetemplatepath)
  - [3.2. SignatureIniPath](#32-signatureinipath)
  - [3.3. ReplacementVariableConfigFile](#33-replacementvariableconfigfile)
  - [3.4. GraphConfigFile](#34-graphconfigfile)
  - [3.5. TrustsToCheckForGroups](#35-truststocheckforgroups)
  - [3.6. IncludeMailboxForestDomainLocalGroups](#36-includemailboxforestdomainlocalgroups)
  - [3.7. DeleteUserCreatedSignatures](#37-deleteusercreatedsignatures)
  - [3.8. DeleteScriptCreatedSignaturesWithoutTemplate](#38-deletescriptcreatedsignatureswithouttemplate)
  - [3.9. SetCurrentUserOutlookWebSignature](#39-setcurrentuseroutlookwebsignature)
  - [3.10. SetCurrentUserOOFMessage](#310-setcurrentuseroofmessage)
  - [3.11. OOFTemplatePath](#311-ooftemplatepath)
  - [3.12. OOFIniPath](#312-oofinipath)
  - [3.13. AdditionalSignaturePath](#313-additionalsignaturepath)
  - [3.14. UseHtmTemplates](#314-usehtmtemplates)
  - [3.15. SimulateUser](#315-simulateuser)
  - [3.16. SimulateMailboxes](#316-simulatemailboxes)
  - [3.17. SimulateTime](#317-simulatetime)
  - [3.18. GraphCredentialFile](#318-graphcredentialfile)
  - [3.19. GraphOnly](#319-graphonly)
  - [3.20. CreateRtfSignatures](#320-creatertfsignatures)
  - [3.21. CreateTxtSignatures](#321-createtxtsignatures)
  - [3.22. EmbedImagesInHtml](#322-embedimagesinhtml)
  - [3.23. DocxHighResImageConversion](#323-docxhighresimageconversion)
  - [3.24. SignaturesForAutomappedAndAdditionalMailboxes](#324-signaturesforautomappedandadditionalmailboxes)
  - [3.25. DisableRoamingSignatures](#325-disableroamingsignatures)
  - [3.26. MirrorLocalSignaturesToCloud](#326-mirrorlocalsignaturestocloud)
  - [3.27. BenefactorCircleId](#327-benefactorcircleid)
  - [3.28. BenefactorCircleLicenceFile](#328-benefactorcirclelicencefile)
- [4. Outlook signature path](#4-outlook-signature-path)
- [5. Mailboxes](#5-mailboxes)
- [6. Group membership](#6-group-membership)
- [7. Removing old signatures](#7-removing-old-signatures)
- [8. Error handling](#8-error-handling)
- [9. Run script while Outlook is running](#9-run-script-while-outlook-is-running)
- [10. Signature and OOF file format](#10-signature-and-oof-file-format)
  - [10.1. Relation between template file name and Outlook signature name](#101-relation-between-template-file-name-and-outlook-signature-name)
  - [10.2. Proposed template and signature naming convention](#102-proposed-template-and-signature-naming-convention)
- [11. Template tags and ini files](#11-template-tags-and-ini-files)
  - [11.1. Allowed tags](#111-allowed-tags)
  - [11.2. How to work with ini files](#112-how-to-work-with-ini-files)
- [12. Signature and OOF application order](#12-signature-and-oof-application-order)
- [13. Variable replacement](#13-variable-replacement)
  - [13.1. Photos from Active Directory](#131-photos-from-active-directory)
    - [13.1.1. When using DOCX template files](#1311-when-using-docx-template-files)
    - [13.1.2. When using HTM template files](#1312-when-using-htm-template-files)
    - [13.1.3. Common behavior](#1313-common-behavior)
  - [13.2. Delete images when attribute is empty, variable content based on group membership](#132-delete-images-when-attribute-is-empty-variable-content-based-on-group-membership)
- [14. Outlook Web](#14-outlook-web)
- [15. Hybrid and cloud-only support](#15-hybrid-and-cloud-only-support)
  - [15.1. Basic Configuration](#151-basic-configuration)
  - [15.2. Advanced Configuration](#152-advanced-configuration)
  - [15.3. Authentication](#153-authentication)
- [16. Simulation mode](#16-simulation-mode)
- [17. FAQ](#17-faq)
  - [17.1. Where can I find the changelog?](#171-where-can-i-find-the-changelog)
  - [17.2. How can I contribute, propose a new feature or file a bug?](#172-how-can-i-contribute-propose-a-new-feature-or-file-a-bug)
  - [17.3. How is the account of a mailbox identified?](#173-how-is-the-account-of-a-mailbox-identified)
  - [17.4. How is the personal mailbox of the currently logged-in user identified?](#174-how-is-the-personal-mailbox-of-the-currently-logged-in-user-identified)
  - [17.5. Which ports are required?](#175-which-ports-are-required)
  - [17.6. Why is Out of Office abbreviated OOF and not OOO?](#176-why-is-out-of-office-abbreviated-oof-and-not-ooo)
  - [17.7. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.](#177-should-i-use-docx-or-htm-as-file-format-for-templates-signatures-in-outlook-sometimes-look-different-than-my-templates)
  - [17.8. How can I log the script output?](#178-how-can-i-log-the-script-output)
  - [17.9. How can I get more script output for troubleshooting?](#179-how-can-i-get-more-script-output-for-troubleshooting)
  - [17.10. Can multiple script instances run in parallel?](#1710-can-multiple-script-instances-run-in-parallel)
  - [17.11. How do I start the script from the command line or a scheduled task?](#1711-how-do-i-start-the-script-from-the-command-line-or-a-scheduled-task)
    - [17.11.1. Start Set-OutlookSignatures in hidden/invisible mode](#17111-start-set-outlooksignatures-in-hiddeninvisible-mode)
  - [17.12. How to create a shortcut to the script with parameters?](#1712-how-to-create-a-shortcut-to-the-script-with-parameters)
  - [17.13. What is the recommended approach for implementing the software?](#1713-what-is-the-recommended-approach-for-implementing-the-software)
  - [17.14. What is the recommended approach for custom configuration files?](#1714-what-is-the-recommended-approach-for-custom-configuration-files)
  - [17.15. Isn't a plural noun in the script name against PowerShell best practices?](#1715-isnt-a-plural-noun-in-the-script-name-against-powershell-best-practices)
  - [17.16. The script hangs at HTM/RTF export, Word shows a security warning!?](#1716-the-script-hangs-at-htmrtf-export-word-shows-a-security-warning)
  - [17.17. How to avoid blank lines when replacement variables return an empty string?](#1717-how-to-avoid-blank-lines-when-replacement-variables-return-an-empty-string)
  - [17.18. Is there a roadmap for future versions?](#1718-is-there-a-roadmap-for-future-versions)
  - [17.19. How to deploy signatures for "Send As", "Send On Behalf" etc.?](#1719-how-to-deploy-signatures-for-send-as-send-on-behalf-etc)
  - [17.20. Can I centrally manage and deploy Outook stationery with this script?](#1720-can-i-centrally-manage-and-deploy-outook-stationery-with-this-script)
  - [17.21. Why is dynamic group membership not considered on premises?](#1721-why-is-dynamic-group-membership-not-considered-on-premises)
    - [17.21.1. Graph](#17211-graph)
    - [17.21.2. Active Directory on premises](#17212-active-directory-on-premises)
  - [17.22. Why is no admin or user GUI available?](#1722-why-is-no-admin-or-user-gui-available)
  - [17.23. What if a user has no Outlook profile or is prohibited from starting Outlook?](#1723-what-if-a-user-has-no-outlook-profile-or-is-prohibited-from-starting-outlook)
  - [17.24. What if Outlook is not installed at all?](#1724-what-if-outlook-is-not-installed-at-all)
  - [17.25. What about the roaming signatures feature in Exchange Online?](#1725-what-about-the-roaming-signatures-feature-in-exchange-online)
    - [17.25.1. Please be aware of the following problem](#17251-please-be-aware-of-the-following-problem)
  - [17.26. Why does the text color of my signature change sometimes?](#1726-why-does-the-text-color-of-my-signature-change-sometimes)
  - [17.27. How to make Set-OutlookSignatures work with Microsoft Information Protection?](#1727-how-to-make-set-outlooksignatures-work-with-microsoft-information-protection)
  - [17.28. Images in signatures have a different size than in templates, or a black background](#1728-images-in-signatures-have-a-different-size-than-in-templates-or-a-black-background)
  
# 1. Requirements  
Requires Outlook and Word, at least version 2010.  
The script must run in the security context of the currently logged-in user.

The script must run in PowerShell Full Language mode. Constrained Language mode is not supported, as some features such as BASE64 conversions are not available in this mode or require very slow workarounds.

If you use AppLocker or a comparable solution, you may need to digitally sign the PowerShell 'Set-OutlokSignatures.ps1'. It is usually not necessary to sign the variable replacement configuration files, e. g. '.\config\default replacement variables.ps1'.  
There are locked down environments, where all files matching the patterns `*.ps*1` and `*.dll` need to be digitially signed with a trusted certificate. 

**Thanks to our partnership with [ExplicIT Consulting](https://explicitconsulting.at), Set-OutlookSignatures and its components are digitally signed with an Extended Validation (EV) Code Signing Certificate (which is the highest code signing standard available).  
This is not only available for Benefactor Circle members, but also the Free and Open Source core version is code signed.**

Don't forget to unblock at least 'Set-OutlookSignatures.ps1' after extracting them from the downloaded ZIP file. You can use the PowerShell commandlet 'Unblock-File' for this.

The paths to the template files (SignatureTemplatePath, OOFTemplatePath) must be accessible by the currently logged-in user. The template files must be at least readable for the currently logged-in user.

In cloud environments, you need to register Set-OutlookSignatures as app and provide admin consent for the required permissions. See '.\config\default graph config.ps1' for details.
# 2. Quick Start Guide
If you already use Set-OutlookSignatures and plan to update to a newer version, start with the CHANGELOG document.

If you are new to Set-OutlookSignatures, start with the README file, which is the document presented right at https://github.com/GruberMarkus/Set-OutlookSignatures - and which you are obviously reading right now.

Read the following chapters in order:
1. Features: Gives you an overview of what Set-OutlookSignatures can do for you.<br>This is the chapter right at the beginning of this document, before the Table of Contents.
2. Chapter 1 (Requirements): Describes the very basic requirements that need to be prepared to run Set-OutlookSignatures in your environment. 
3. Chapter 2 (Quick Start Guide): You are right here.
4. Chapter 3 (Parameters): Describes how the behavior of Set-OutlookSignatures can be modified.<br>This gives you a deeper understanding of the features, but also answers how you can change the behavior of Set-OutlookSignatures.
5. Should you not be familiar with basic usage of PowerShell, i.e. starting PowerShell and running existing scripts, please ask your IT department for support. You can start learning the basics [here](https://learn.microsoft.com/en-us/powershell/scripting/learn/ps101/01-getting-started?view=powershell-5.1).

You now have a good theoretical overview.<br>If you want to know more before you start with the practical implementation, just read through the rest of the README file.

**Set-OutlookSignatures is very well documented, which inevitably brings with it a lot of content.  
If you are looking for someone with experience who can quickly train you and assist with evaluation, planning, implementation and ongoing operations: Our partner [ExplicIT Consulting](https://explicitconsulting.at) offers professional support and also the [Benefactor Circle add-on](https://explicitconsulting.at/open-source/set-outlooksignatures) with enhanced features.**

To start with the practical implementation:
1. Log on to a Windows client with a test user having an Exchange mailbox, and configure Outlook
2. Download Set-OutlookSignatures and extract the archive to a local folder
3. Unblock the file 'Set-OutlookSignatures.ps1'.<br>You can use the PowerShell commandlet 'Unblock-File' for this, or right-click the file in File Explorer, select Properties and check 'Unblock'.
4. If you use AppLocker or a comparable solution, you may need to digitally sign the PowerShell 'Set-OutlokSignatures.ps1'.  
It is usually not necessary to sign the variable replacement configuration files, e. g. '.\config\default replacement variables.ps1'.  
There are locked down environments, where all files matching the patterns `*.ps*1` and `*.dll` need to be digitially signed with a trusted certificate. 
5. Now it is time to run Set-OutlookSignatures for the first time
   - Mailbox is in Exchange on-prem and the logged-in user has access to the on-prem Active Directory: Run 'Set-OutlookSignatures.ps1' in PowerShell.<br>For best results, don't run the script by double clicking it in File Explorer, or via right-click and 'Run'. Instead, run the following command:
      ```
      powershell.exe -noexit -file "c:\test\Set-OutlookSignatures.ps1" # adopt the file path as needed
      ```
   - Mailbox is in Exchange Online and/or the logged-in user has no access to the on-prem Active Directory and/or your environment is cloud only: You need to register an Entra ID/Azure AD application first, because Set-OutlookSignatures needs permissions to access the Graph API.
     1. An Entra ID/Azure AD administrator needs to register Set-OutlookSignatures as app and provide admin consent for the required permissions. See the file '.\config\default graph config.ps1' for details.
     2. Run Set-OutlookSignatures
        - Cloud only, or hybrid without synced Exchange attributes (mail, legacyExchangeDN, msExchRecipientTypeDetails, msExchMailboxGuid, proxyAddresses):
          ```
          powershell.exe -noexit -file "c:\test\Set-OutlookSignatures.ps1" -GraphOnly true # adopt the file path as needed
          ```
        - Hybrid with synced Exchange attributes (mail, legacyExchangeDN, msExchRecipientTypeDetails, msExchMailboxGuid, proxyAddresses):
          ```
          powershell.exe -noexit -file "c:\test\Set-OutlookSignatures.ps1" # adopt the file path as needed
          ```

Set-OutlookSignatures now runs using default settings and sample templates.<br>Because of the '-exit' parameter, the window hosting Set-OutlookSignatures will not close after the script completed. This is helpful for debugging and learning.

Watch the output for errors and warnings, displayed in red or yellow when run in the PowerShell console.
- If there are errors or warnings:
  1. Read the messages carefully as they often contain hints on how to resolve the issue.  
  2. Check if the README file contains a hint.
  3. Check if someone has already reported the problem as and issue on [GitHub](https://github.com/GruberMarkus/Set-OutlookSignatures/issues?q=is%3Aissue), and create a new one if you can't find any hint on how to solve it.
- If there are no errors, switch to Outlook and have a look at the newly created signatures.

When everything runs fine with default settings, it is time to start customizing the script behavior to your needs:
- Create a folder with your own template files and signature configuration file.
  - Start with DOCX templates. See `Should I use .docx or .htm as file format for templates?` in this document for details.
  - See the following chapters in this document for instructions
    - Signature and OOF file format
    - Signature template file naming
    - Template tags and ini files
  - Make sure to pass the parameters `SignatureTemplatePath`, `SignatureIniPath`, `OOFTemplatePath` and `OOFInipath` to Set-OutlookSignatures
- Adopt other parameters you may find useful, or start experimenting with simulation mode.<br>The feature list and the parameter documentation show what's possible.
<br>
<br>
<p>  
It is strongly recommended to not change any Set-OutlookSignatures files and keep them as they are. If you consequently work with script parameters and keep customized configuration files in a separate folder, upgrading to a new version is basically just a file copy operation (drop-in replacement).

Regarding configuration files: Besides the template configuration files for signatures and OOF messages, there are the Graph configuration file and the replacement variable configuration file.  
It is rarely needed to change the configuration within these files.<br>The configuration files themselves contain specific information on how to use them.<br>The configuration files are referenced in the documentation whenever there is a need or option to change them.

**[Benefactor Circle members](https://explicitconsulting.at/open-source/set-outlooksignatures) have access to a document covering the organizational aspects of introducing Set-OutlookSignatures.**  
The content is based on real life experiences implementing the script in multi-client environments with a five-digit number of mailboxes.  
It contains proven procedures and recommendations for product managers, architects, operations managers, account managers and e-mail and client administrators. It is suited for service providers as well as for clients.  
It covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.
# 3. Parameters  
## 3.1. SignatureTemplatePath  
The parameter SignatureTemplatePath tells the script where signature template files are stored.

Local and remote paths are supported. Local paths can be absolute (`C:\Signature templates`) or relative to the script path (`.\templates\Signatures`).

WebDAV paths are supported (https only): `https://server.domain/SignatureSite/SignatureTemplates` or `\\server.domain@SSL\SignatureSite\SignatureTemplates`

The currently logged-in user needs at least read access to the path.

Default value: `.\templates\Signatures DOCX`  
## 3.2. SignatureIniPath
Template tags are placed in an ini file.

The file must be UTF8 encoded.

See '.\templates\Signatures DOCX\_Signatures.ini' for a sample file with further explanations.

Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures')

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'

The currently logged-in user needs at least read access to the path

Default value: `.\templates\Signatures DOCX\_Signatures.ini`
## 3.3. ReplacementVariableConfigFile  
The parameter ReplacementVariableConfigFile tells the script where the file defining replacement variables is located.

The file must be UTF8 encoded.

Local and remote paths are supported. Local paths can be absolute (`C:\config\default replacement variables.ps1`) or relative to the script path (`.\config\default replacement variables.ps1`).

WebDAV paths are supported (https only): `https://server.domain/SignatureSite/config/default replacement variables.ps1` or `\\server.domain@SSL\SignatureSite\config\default replacement variables.ps1`

The currently logged-in user needs at least read access to the file.

Default value: `.\config\default replacement variables.ps1`  
## 3.4. GraphConfigFile
The parameter GraphConfigFile tells the script where the file defining Graph connection and configuration options is located.

The file must be UTF8 encoded.

Local and remote paths are supported. Local paths can be absolute (`C:\config\default graph config.ps1`) or relative to the script path (`.\config\default graph config.ps1`).

WebDAV paths are supported (https only): `https://server.domain/SignatureSite/config/default graph config.ps1` or `\\server.domain@SSL\SignatureSite\config\default graph config.ps1`

The currently logged-in user needs at least read access to the file.

Default value: `.\config\default graph config.ps1`  
## 3.5. TrustsToCheckForGroups  
List of domains to check for group membership.

If the first entry in the list is '*', all outgoing and bidirectional trusts in the current user's forest are considered.

If a string starts with a minus or dash ('-domain-a.local'), the domain after the dash or minus is removed from the list (no wildcards allowed).

All domains belonging to the Active Directory forest of the currently logged-in user are always considered, but specific domains can be removed (`*', '-childA1.childA.user.forest`).

When a cross-forest trust is detected by the '*' option, all domains belonging to the trusted forest are considered but specific domains can be removed (`*', '-childX.trusted.forest`).

Default value: '*'
## 3.6. IncludeMailboxForestDomainLocalGroups
Shall the script consider group membership in domain local groups in the mailbox's AD forest?

Per default, membership in domain local groups in the mailbox's forest is not considered as the required LDAP queries are slow and domain local groups are usually not used in Exchange.

Domain local groups across trusts behave differently, they are always considered as soon as the trusted domain/forest is included in TrustsToCheckForGroups.

Default value: $false
## 3.7. DeleteUserCreatedSignatures  
Shall the script delete signatures which were created by the user itself?

Default value: `$false`
## 3.8. DeleteScriptCreatedSignaturesWithoutTemplate
Shall the script delete signatures which were created by the script before but are no longer available as template?

Default value: `$true`
## 3.9. SetCurrentUserOutlookWebSignature  
Shall the script set the Outlook Web signature of the currently logged-in user?

If the parameter is set to `$true` and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. This way, the script can be used in environments where only Outlook Web is used. 

Default value: `$true`  
## 3.10. SetCurrentUserOOFMessage  
Shall the script set the Out of Office (OOF) auto reply message of the currently logged-in user?

If the parameter is set to `$true` and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. This way, the script can be used in environments where only Outlook Web is used. 

Default value: `$true`  
## 3.11. OOFTemplatePath  
Path to centrally managed Out of Office (OOF) auto reply templates.

Local and remote paths are supported.

Local paths can be absolute (`C:\OOF templates`) or relative to the script path (`.\templates\Out of Office`).

WebDAV paths are supported (https only): `https://server.domain/SignatureSite/OOFTemplates` or `\\server.domain@SSL\SignatureSite\OOFTemplates`

The currently logged-in user needs at least read access to the path.

Default value: `.\templates\Out of Office DOCX`
## 3.12. OOFIniPath
Template tags are placed in an ini file.

The file must be UTF8 encoded.

See '.\templates\Out of Office DOCX\_OOF.ini' for a sample file with further explanations.

Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures')

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'

The currently logged-in user needs at least read access to the path

Default value: `.\templates\Out of Office DOCX\_OOF.ini`
## 3.13. AdditionalSignaturePath  
An additional path that the signatures shall be copied to.  
Ideally, this path is available on all devices of the user, for example via Microsoft OneDrive or Nextcloud.

This way, the user can easily copy-paste his preferred preconfigured signature for use in an e-mail app not supported by this script, such as Microsoft Outlook Mobile, Apple Mail, Google Gmail or Samsung Email.

Local and remote paths are supported.

Local paths can be absolute (`C:\Outlook signatures`) or relative to the script path (`.\Outlook signatures`).

WebDAV paths are supported (https only): `https://server.domain/User/Outlook signatures` or `\\server.domain@SSL\User\Outlook signatures`

The currently logged-in user needs at least write access to the path.

If the folder or folder structure does not exist, it is created.

Default value: `"$([environment]::GetFolderPath("MyDocuments"))\Outlook signatures"`  
## 3.14. UseHtmTemplates  
With this parameter, the script searches for templates with the extension .htm instead of .docx.

Templates in .htm format must be UTF8 encoded and the charset must be set to UTF8 (`<META content="text/html; charset=utf-8">`).

Each format has advantages and disadvantages, please see "[13.5. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.](#135-should-i-use-docx-or-htm-as-file-format-for-templates-signatures-in-outlook-sometimes-look-different-than-my-templates)" for a quick overview.

Default value: `$false`  
## 3.15. SimulateUser  
SimulateUser is a mandatory parameter for simulation mode. This value replaces the currently logged-in user.

Use a logon name in the format 'Domain\User' or a Universal Principal Name (UPN, looks like an e-mail-address, but is not neecessarily one).

See "[13. Simulation mode](#13-simulation-mode)" for details.  
## 3.16. SimulateMailboxes  
SimulateMailboxes is optional for simulation mode, although highly recommended. It is a comma separated list of e-mail addresses replacing the list of mailboxes otherwise gathered from the registry.
## 3.17. SimulateTime  
SimulateTime is optional for simulation mode.

Use a certain timestamp for simulation mode. This allows you to simulate time-based templates.

Format: yyyyMMddHHmm (yyyy = year, MM = two-digit month, dd = two-digit day, HH = two-digit hour (0..24), mm = two-digit minute), local time

Default value: $null
## 3.18. GraphCredentialFile
Path to file containing Graph credential which should be used as alternative to other token acquisition methods.

Makes only sense in combination with `.\sample code\SimulateAndDeploy.ps1`, do not use this parameter for other scenarios.

See `.\sample code\SimulateAndDeploy.ps1` for an example how to create this file.

Default value: `$null`  
## 3.19. GraphOnly
Try to connect to Microsoft Graph only, ignoring any local Active Directory.

The default behavior is to try Active Directory first and fall back to Graph.

Default value: `$false`
## 3.20. CreateRtfSignatures
Should signatures be created in RTF format?

Default value: `$false`
## 3.21. CreateTxtSignatures
Should signatures be created in TXT format?

Default value: `$true`
## 3.22. EmbedImagesInHtml
Should images be embedded into HTML files?

Outlook 2016 and newer can handle images embedded directly into an HTML file as BASE64 string (`<img src="data:image/[...]"`).

Outlook 2013 and earlier can't handle these embedded images when composing HTML e-mails (there is no problem receiving such e-mails, or when composing RTF or TXT e-mails).

When setting EmbedImagesInHtml to `$false`, consider setting the Outlook registry value "Send Pictures With Document" to 1 to ensure that images are sent to the recipient (see https://support.microsoft.com/en-us/topic/inline-images-may-display-as-a-red-x-in-outlook-704ae8b5-b9b6-d784-2bdf-ffd96050dfd6 for details). Set-OutlookSignatures does this automatically for the currently logged-in user, but it may be overridden by other scripts or group policies.

Default value: `$false`
## 3.23. DocxHighResImageConversion
Enables or disables high resolution images in HTML signatures.

When enabled, this parameter uses a workaround to overcome a Word limitation that results in low resolution images when converting to HTML. The price for high resolution images in HTML signatures are more time needed for document conversion and signature files requiring more storage space.

Disabling this feature speeds up DOCX to HTML conversion, and HTML signatures require less storage space - at the cost of lower resolution images.

Contrary to conversion to HTML, conversion to RTF always results in high resolution images.

Default value: $true
## 3.24. SignaturesForAutomappedAndAdditionalMailboxes
Deploy signatures for automapped mailboxes and additional mailboxes.

Signatures can be deployed for these mailboxes, but not set as default signature due to technical restrictions in Outlook.

Default value: `$true`
## 3.25. DisableRoamingSignatures
Disable roaming signatures.

Only sets HKCU registry key, does not override configuration set by group policy.

Possible values: $null, $true, $false

Default value: $true
## 3.26. MirrorLocalSignaturesToCloud
Should local signatures be uploaded as roaming signature for the current user?

Possible for Exchange Online mailbox of currently logged-in user.

Prerequisites:
- Script parameter `SetCurrentUserOutlookWebSignature` is set to `true`
- Mailbox is the mailbox of the currently logged-in user and is hosted in Exchange Online
- Script parameter `MirrorLocalSignaturesToCloud` is set to `true`
  - This is independent from the Outlook registry key `DisableRoamingSignatures`, but: It is recommended to disable the Outlook roaming feature by setting the script parameter `DisableRoamingSignatures` to `true` when MirrorLocalSignaturesToCloud is enabled. This avoids Outlook showing synced signatures twice (one as local one, one as roaming one).

Please note:
- As there is no Microsoft official API, this feature is experimental, and you use it on your own risk.
- This feature does not work in simulation mode, because the user running the simulation does not have access to the signatures stored in another mailbox

The process is very simple and straight forward. Set-OutlookSignatures goes through the following steps:
1. Check if all prerequisites are met
2. Download all signatures stored in the Exchange Online mailbox of the currently logged-in user
  - This mimics Outlook's behavior: Roaming signatures are only manipulated in the cloud and then downloaded from there.
  - Local signatures with identical names are overwritten with the content of the roaming signature
3. Go through standard template and signature processing
  - Loop through the templates and their configuration, and convert them to signatures
  - Set default signatures for replies and forwards
  - If configured, delete signatures created by the user
  - If configured, delete signatures created earlier by Set-OutlookSignatures but now no longer have a corresponding central configuration
4. Delete all signatures in the cloud and upload all locally stored signatures to the user's personal mailbox as roaming signatures


This process is very simple and straight forward. It basically just mirrors the locally available signatures as roaming signatures, so there may be too many signatures available in the cloud - in this case, having too many signatures is better than missing some.

Another advantage of this solution is that it makes roaming signatures available in Outlook versions that do not officially support them.

Default value: $false

## 3.27. BenefactorCircleId
The Benefactor Circle member Id matching your licence file, which unlocks exclusive features.

Default value: ''
## 3.28. BenefactorCircleLicenceFile
The Benefactor Circle licence file matching your member Id, which unlocks exclusive features.

Default value: ''
# 4. Outlook signature path  
The Outlook signature path is retrieved from the users registry, so the script is language independent.

The registry setting does not allow for absolute paths, only for paths relative to `%APPDATA%\Microsoft`.

If the relative path set in the registry would be a valid path but does not exist, the script creates it.  
# 5. Mailboxes  
The script only considers primary mailboxes per default, these are mailboxes added as separate accounts.

This is the same way Outlook handles mailboxes from a signature perspective: Outlook can not handle signatures for non-primary mailboxes (added via "Open these additional mailboxes").

Setting the parameter `SignaturesForAutomappedAndAdditionalMailboxes` to `true` allows the script to detect automapped and additional mailboxes. Signatures can be deployed for these types of mailboxes, but they can not be set as default signatures due to technical restrictions in Outlook.

The script is created for Exchange environments. Non-Exchange mailboxes can not have OOF messages or group signatures, but common and mailbox specific signatures.  
# 6. Group membership  
Per default, the script considers all static security and distribution groups of group scopes global and universal the currently processed mailbox belongs to.

Group membership is evaluated against the whole Active Directory forest of the mailbox, and against all trusted domains (and their subdomains) the user has access to.

In Exchange resource forest scenarios with linked mailboxes, the group membership of the linked account (as populated in msExchMasterAccountSID) is not considered, only the group membership of the actual mailbox.

Group membership from Active Directory on-prem is retrieved by combining two queries:
- Security and distribution groups, are queried via the tokenGroupsGlobalAndUniversal attribute. Querying this attribute is very fast, resource saving on client and server, and also considers sIDHistory. This query only includes groups with the global or universal scope type.
- Group membership across trusts is always considered, as soon as the trusted domain/forest is included in TrustsToCheckForGroups. Cross-trust group membership is retrieved with an optimized LDAP query, considering the sID and sIDHistory of the group memberships retrieved in the steps before. This query only includes groups with the domain local scope type, as this is the only group type that allows for group membership across trusts.
- Only static groups are considered. Please see the FAQ section for detailed information why dynamic groups are not included in group membership queries on-prem.
- Per default, the mailbox's own forest is not checked for membership in domain local groups, no matter if of type security or distribution. This is because querying for domain local groups takes time, as every domain local group domain in the forest has to be queried for membership. Also, domain local groups are usually not used when granting permissions in Exchange. You can enable searching for domain local groups in the mailbox's forest by setting the parameter `IncludeMailboxForestDomainLocalGroups` to `$true`.


When no Active Directory connection is available, Microsoft Graph is queried for transitive group membership. This query includes security and distribution groups.  
In Microsoft Graph, membership in dynamic groups is considered, too.

# 7. Removing old signatures  
The script always deletes signatures which were deployed by the script earlier, but are no longer available in the central repository. The script marks each processed signature with a specific HTML tag, which enables this cleaning feature.

Signatures created manually by the user are not deleted by default, this behavior can be changed with the DeleteUserCreatedSignatures parameter.  
# 8. Error handling  
Error handling is implemented rudimentarily.  
# 9. Run script while Outlook is running  
Outlook and the script can run simultaneously.

New and changed signatures can be used instantly in Outlook.

Changing which signature is to be used as default signature for new e-mails or for replies and forwards requires restarting Outlook.   
# 10. Signature and OOF file format  
Only Word files with the extension .docx and HTML files with the extension .htm are supported as signature and OOF template files.  
## 10.1. Relation between template file name and Outlook signature name  
The name of the signature template file without extension is the name of the signature in Outlook.
Example: The template "Test signature.docx" will create a signature named "Test signature" in Outlook.

This can be overridden in the ini file with the 'OutlookSignatureName' parameter.
Example: The template "Test signature.htm" with the following ini file configuration will create a signature named "Test signature, do not use".
```
[Test signature.htm]
OutlookSignatureName = Test signature, do not use
```
## 10.2. Proposed template and signature naming convention
To make life easier for template maintainers and for users, a consistent template and signature naming convention should be used.

There are multiple approaches, with the following one gaining popularity: `<Company> <internal/external> <Language> <formal/informal> <additional info>`

Let's break down the components:
- Company: Useful when your users work with multiple company or brand names.
- Internal/External: Usually abbreviated as int and ext. Show if a signature is intended for use with a purely internal recipient audience, or if an external audience is involved.
- Language: Usually abbreviated to a two-letter code, such as AT for Austria. This way, you can handle multi-language signatures.
- Formal/informal: Usually abbreviated as frml and infrml. Allows you to deploy signatures with a certain formality in the salutation of the signature.
- Additional info: Typically used to identify signatures for shares mailboxes or in delegate scenarios.

Example signature names for a user having access to his own mailbox and the office mailbox:
- CompA ext DE frml
- CompA ext DE frml office@
- CompA ext DE infrml
- CompA ext DE infrml office@
- CompA ext EN frml
- CompA ext EN frml office@
- CompA ext EN infrml
- CompA ext EN infrml office@
- CompA int EN infrml
- CompA int EN infrml office@

For the user, the selection process may look complicated at first sight, but is actually quite natural and fast:
- Example A: Informal German mail sent to externals from own mailbox
  1. "I act in the name of company CompA" -> "CompA"
  2. "The mail has at least one external recipient" -> "CompA ext"
  3. "The mail is written in German language" -> "CompA ext DE"
  4. "The tone is informal" -> "CompA ext DE infrml"
  5. "I send from my own mailbox" -> no change, use "CompA ext DE infrml"
- Example B: Formal English mail sent to externals from office@
  1. "I act in the name of company CompA" -> "CompA"
  2. "The mail has at least one external recipient" -> "CompA ext"
  3. "The mail is written in English language" -> "CompA ext EN"
  4. "The tone is formal" -> "CompA ext EN frml"
  5. "I send from the office mailbox" -> "CompA ext EN frml office@"
- Example C: Internal English mail from own mailbox
  1. "I act in the name of company CompA" -> "CompA"
  2. "The mail has only internal recipients" -> "CompA int"
  3. "The mail is written in English language" -> "CompA int EN"
  4. "The tone is informal" -> "CompA int EN infrml"
  5. "I send from my own mailbox" -> "CompA int EN infrml"

Don't forget: You can use one and the same template for different signature names. In the example above, the template might not be named `CompA ext EN frml office@.docx`, but `CompA ext EN frml shared@.docx` and be used multiple times in the ini file:
```
# office@example.com
[CompA ext EN frml shared@.docx]
office@example.com
OutlookSignatureName = CompA ext EN frml office@
DefaultNew

# marketing@example.com
[CompA ext EN frml shared@.docx]
marketing@example.com
OutlookSignatureName = CompA ext EN frml marketing@
DefaultNew
```
# 11. Template tags and ini files
Tags define properties for templates, such as
- time ranges during which a template shall be applied or not applied
- groups whose direct or indirect members are allowed or denied application of a template
- specific e-mail addresses which are are allowed or denied application of a template
- specific replacement variables which allow or deny application of a template
- an Outlook signature name that is different from the file name of the template
- if a signature template shall be set as default signature for new e-mails or as default signature for replies and forwards
- if a OOF template shall be set as internal or external message

There are additional tags which are not template specific, but change the behavior of Set-OutlookSignatures:
- specific sort order for templates (ascending, descending, as listed in the file)
- specific sort culture used for sorting ascendingly or descendingly (de-AT or en-US, for example)

If you want to give template creators control over the ini file, place it in the same folder as the templates.

Tags are case insensitive.
## 11.1. Allowed tags
- `<yyyyMMddHHmm-yyyyMMddHHmm>`, `-:<yyyyMMddHHmm-yyyyMMddHHmm>`
  - Make this template valid only during the specific time range (`yyyy` = year, `MM` = month, `dd` = day, `HH` = hour (00-24), `mm` = minute).
  - The `-:` prefix makes this template invalid during the specified time range.
  - Examples: `202112150000-202112262359` for the 2021 Christmas season, `-:202202010000-202202282359` for a deny in February 2022
  - If the script does not run after a template has expired, the template is still available on the client and can be used.
  - Time ranges are interpreted as local time per default, which means times depend on the user or client configuration. If you do not want to use local times, but global times just add 'Z' as time zone. For example: `202112150000Z-202112262359Z`
- `<NetBiosDomain> <GroupSamAccountName>`, `<NetBiosDomain> <Display name of Group>`, `-:<NetBiosDomain> <GroupSamAccountName>`, `-:<NetBiosDomain> <Display name of Group>`
  - Make this template specific for an Outlook mailbox being a direct or indirect member of this group or distribution list
  - The `-:` prefix makes this template invalid for the specified group.
  - Examples: `EXAMPLE Domain Users`, `-:Example GroupA`  
  - Groups must be available in Active Directory. Groups like `Everyone` and `Authenticated Users` only exist locally, not in Active Directory
  - This tag supports alternative formats, which are of special interest if you are in a cloud only or hybrid environmonent:
    - `<NetBiosDomain> <GroupSamAccountName>` and `<NetBiosDomain> <Group DisplayName>` can be queried from Microsoft Graph if the groups are synced between on-prem and the cloud. SamAccountName is queried before DisplayName. Use these formats when your environment is hybrid or on premises only.
    - `AzureAD <e-mail-address-of-group@example.com>`, `AzureAD <GroupMailNickname>`, `AzureAD <GroupDisplayName>`, `EntraID <e-mail-address-of-group@example.com>`, `EntraID <GroupMailNickname>`, `EntraID <GroupDisplayName>` do not work with a local Active Directory, only with Microsoft Graph. They are queried in the order given. 'EntraID' and 'AzureAD' are the literal, case-insensitive strings 'EntraID' and 'AzureAD', not a variable. Use these formats when you are in a cloud only environment.
  - '<NetBiosDomain>' and '<EXAMPLE>' are just examples. You need to replace them with the actual NetBios domain name of the Active Director domain containing the group.
  - 'EntraID' and 'AzureAD' are not examples. If you want to assign a template to a group stored in Azure Active Directory, you have to use 'EntraID' or 'AzureAD' as domain name.
  - When multiple groups are defined, membership in a single group is sufficient to be assigned the template - it is not required to be a member of all the defined groups.  
- `CURRENTUSER:<NetBiosDomain> <GroupSamAccountName>`, `CURRENTUSER:<NetBiosDomain> <Display name of Group>`, `-CURRENTUSER:<NetBiosDomain> <GroupSamAccountName>`, `-CURRENTUSER:<NetBiosDomain> <Display name of Group>`, `CURRENTUSER:AzureAD <e-mail-address-of-group@example.com>`, `CURRENTUSER:AzureAD <GroupMailNickname>`, `CURRENTUSER:AzureAD <GroupDisplayName>`, `-CURRENTUSER:AzureAD <e-mail-address-of-group@example.com>`, `-CURRENTUSER:AzureAD <GroupMailNickname>`, `-CURRENTUSER:AzureAD <GroupDisplayName>`, `CURRENTUSER:EntraID <e-mail-address-of-group@example.com>`, `CURRENTUSER:EntraID <GroupMailNickname>`, `CURRENTUSER:EntraID <GroupDisplayName>`, `-CURRENTUSER:EntraID <e-mail-address-of-group@example.com>`, `-CURRENTUSER:EntraID <GroupMailNickname>`, `-CURRENTUSER:EntraID <GroupDisplayName>`
  - Make this template specific for the logged on user if his _personal_ mailbox (which does not need to be in Outlook) is a direct or indirect member of this group or distribution list
  - Example: Assign template to every mailbox, but not if the mailbox of the current user is member of the group EXAMPLE\Group
    ```
    [template.docx]
    -CURRENTUSER:EXAMPLE Group
    ```
- `<SmtpAddress>`, `-:<SmtpAddress>`
  - Make this template specific for the assigned e-mail address (all SMTP addresses of a mailbox are considered, not only the primary one)
  - The `-:` prefix makes this template invalid for the specified e-mail address.
  - Examples: `office@example.com`, `-:test@example.com`
  - The `CURRENTUSER:` and `-CURRENTUSER:` prefixes make this template invalid for the specified e-mail addresses of the current user.  
  Example: Assign template to every mailbox, but not if the personal mailbox of the current user has the e-mail address userX@example.com
  - Useful for delegate or boss-secretary scenarios: "Assign a template to everyone having the boss mailbox userA@example.com in Outlook, but not for UserA itself" is realized like that in the ini file:
    ```
    [delegate template name.docx]
    # Assign the template to everyone having userA@example.com in Outlook
    userA@example.com
    # Do not assign the template to the actual user owning the mailbox userA@example.com
    -CURRENTUSER:userA@example.com
    ```
    You can even only use only one delegate template for your whole company to cover all delegate scenarios. Make sure the template correctly uses `$CurrentUser[...]$` and `$CurrentMailbox[...]$` replacement variables, and then use the template multiple times in the ini file, with different signature names:
    ```
    [Company EN external formal delegate.docx]
    # Assign the template to everyone having userA@example.com in Outlook
    userA@example.com
    # Do not assign the template to the actual user owning the mailbox userA@example.com
    -CURRENTUSER:userA@example.com
    # Use a custom signature name instead of the template file name 
    OutlookSignatureName = Company EN external formal userA@


    [Company EN external formal delegate.docx]
    # Assign the template to everyone having userX@example.com in Outlook
    userX@example.com
    # Do not assign the template to the actual user owning the mailbox userX@example.com
    -CURRENTUSER:userX@example.com
    # Use a custom signature name instead of the template file name 
    OutlookSignatureName = Company EN external formal UserX@
- `<ReplacementVariable>`, `-:<ReplacementVariable>`
  - Make this template specific for the assigned replacement variable
  - The `-:` prefix makes this template invalid for the specified replacement variable.
  - Replacement variable are checked for true or false values. If a replacement variable is not a boolean (true or false) value per se, it is converted to the boolean data type first.
    - Replacement variables that can only hold one value evaluate to false if they contain no value (null, empty) or have the value 0. All other values evaluate to true.
    - Replacement variables that can hold multiple values evaluate to false if they contain no value, or if they contain only one value, which in turn evaluates to false. All other values evaluate to true.
  - Examples:
    - `$CurrentMailboxManagerMail$` (apply if current user has a manager with an e-mail address)
    - `-:$CurrentMailboxManagerMail$` (do not apply if current user has a manager with an e-mail address)
    - A template should only be applied to users which are member of the Marketing group and the Sales group at the same time:
      - Use a custom replacement variable config file, define the custom replacement variable `$CurrentMailbox-IsMemberOf-MarketingAndSales$` and set it to yes if the current user's mailbox is member of the Marketing and the Sales groups at the same time:  
        ```
        @(@('CurrentUser', '$CurrentUser-IsMemberOf-MarketingAndSales$', 'EXAMPLEDOMAIN Marketing', 'EXAMPLEDOMAIN Sales'), @()) | Where-Object { $_ } | Foreach-Object { if ( ((Get-Variable -Name "ADProps$($_[0])" -ValueOnly).GroupsSids -icontains ResolveToSid($_2])) -and ((Get-Variable -Name "ADProps$($_[0])" -ValueOnly).GroupsSids -icontains ResolveToSid($_3])) ) { $ReplaceHash[$_[1]] = 'yes' } else { $ReplaceHash[$_[1]] = $null } }
        ```
      - The template ini configuration then looks like this:
        ```
        [template.docx]
        $CurrentUser-IsMemberOf-MarketingAndSales$
        ```
      - If you want a template only to not be applied to users whose primary mailbox is a of the Marketing group and the Sales group at the same time:
        ```
        [template.docx]
        -:$CurrentUser-IsMemberOf-MarketingAndSales$
        ```
      - Combinations are possible: Only in January 2024, for all members of EXAMPLEDOMAIN\Examplegroup but not for the mailbox example@example.com and not for users whose primary mailbox is a of the Marketing group and the Sales group at the same time:
        ```
        [template.docx]
        202401010000-202401312359
        EXAMPLEDOMAIN Examplegroup
        -:example@example.com
        -:$CurrentUser-IsMemberOf-MarketingAndSales$
        ```
- `writeProtect` (signature template files only)  
    - Write protects the signature files. Works only in Outlook. Modifying the signature in Outlook's signature editor leads to an error on saving, but the signature can still be changed after it has been added to an e-mail.  
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
    ```

<br>Tags can be combined: A template may be assigned to several groups, e-mail addresses and time ranges, be denied for several groups, e-mail adresses and time ranges, be used as default signature for new e-mails and as default signature for replies and forwards - all at the same time. Simple add different tags below a file name, separated by line breaks (each tag needs to be on a separate line).

## 11.2. How to work with ini files
1. Comments
  Comment lines start with '#' or ';'
	Whitespace(s) at the beginning and the end of a line are ignored
  Empty lines are ignored
2. Use the ini files in `.\templates\Signatures DOCX with ini` and `.\templates\Out of Office DOCX with ini` as templates and starting point
3. Put file names with extensions in square brackets  
  Example: `[Company external English formal.docx]`  
  Putting file names in single or double quotes is possible, but not necessary.  
  File names are case insensitive
    `[file a.docx]` is the same as `["File A.docx"]` and `['fILE a.dOCX']`  
  File names not mentioned in this file are not considered, even if they are available in the file system. Set-OutlookSignatures will report files which are in the file system but not mentioned in the current ini, and vice versa.<br>  
  When there are two or more sections for a filename: The keys and values are not combined, each section is considered individually (SortCulture and SortOrder still apply).  
  This can be useful in the following scenario: Multiple shared mailboxes shall use the same template, individualized by using `$CurrentMailbox[...]$` variables. A user can have multiple of these shared mailboxes in his Outlook configuration.
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
    You can even only use only one delegate template for your whole company to cover all delegate scenarios. Make sure the template correctly uses `$CurrentUser[...]$` and `$CurrentMailbox[...]$` replacement variables, and then use the template multiple times in the ini file, with different signature names:
    ```
    [Company EN external formal delegate.docx]
    # Assign the template to everyone having userA@example.com in Outlook
    userA@example.com
    # Do not assign the template to the actual user owning the mailbox userA@example.com
    -CURRENTUSER:userA@example.com
    # Use a custom signature name instead of the template file name 
    OutlookSignatureName = Company EN external formal userA@


    [Company EN external formal delegate.docx]
    # Assign the template to everyone having userX@example.com in Outlook
    userX@example.com
    # Do not assign the template to the actual user owning the mailbox userX@example.com
    -CURRENTUSER:userX@example.com
    # Use a custom signature name instead of the template file name 
    OutlookSignatureName = Company EN external formal UserX@
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
    With this option, you can have different template file names for the same Outlook signature name. Search for "Marketing external English formal" in the sample ini files for an example. Take care of signature group priorities (common, group, e-mail address, replacement variable) and the SortOrder and SortCulture parameters.
5. Remove the tags from the file names in the file system  
Else, the file names in the ini file and the file system do not match, which will result in some templates not being applied.  
It is recommended to create a copy of your template folder for tests.
6. Make the script use the ini file by passing the `SignatureIniPath` and/or `OOFIniPath` parameter
# 12. Signature and OOF application order  
Signatures are applied mailbox for mailbox. The mailbox list is sorted as follows (from highest to lowest priority):
- Mailbox of the currently logged-in user
- Mailboxes from the default Outlook profile, in the sort order shown in Outlook (and not in the order they were added to the Outlook profile)
- Mailboxes from other Outlook profiles. The profiles are sorted alphabetically. Within each profile, the mailboxes are sorted in the order they are shown in Outlook.

For each mailbox, templates are applied in a specific order: Common templates first, group templates second, e-mail address specific templates third, replacement variables last.

Each one of these templates groups can have one or more time range tags assigned. Such a template is only considered if the current system time is within at least one of these time range tags.
- Common templates are templates with either no tag or only `[defaultNew]` and/or `[defaultReplyFwd]` (`[internal]` and/or `[external]` for OOF templates).
- Within these template groups, templates are sorted according to the sort order and sort culture defines in the configuration file.
- Every centrally stored signature template is only applied to the mailbox with the highest priority allowed to use it. This ensures that no mailbox with lower priority can overwrite a signature intended for a higher priority mailbox.

OOF templates are only applied if the Out of Office assistant is currently disabled. If it is currently active or scheduled to be automatically activated in the future, OOF templates are not applied.  
# 13. Variable replacement  
Variables are case sensitive.

Variables are replaced everywhere, including links, QuickTips and alternative text of images.

With this feature, you can not only show e-mail addresses and telephone numbers in the signature and OOF message, but show them as links which open a new e-mail message (`"mailto:"`) or dial the number (`"tel:"`) via a locally installed softphone when clicked.

Custom Active directory attributes are supported as well as custom replacement variables, see `.\config\default replacement variables.ps1` for details.  
Attributes from Microsoft Graph need to be mapped, this is done in `.\config\default graph config.ps1`.

Variables can also be retrieved from other sources than Active Directory by adding custom code to the variable config file.

Per default, `.\config\default replacement variables.ps1` contains the following replacement variables:  
- Currently logged-in user  
    - `$CurrentUserGivenname$`: Given name  
    - `$CurrentUserSurname$`: Surname  
    - `$CurrentUserDepartment$`: Department  
    - `$CurrentUserTitle$`: (Job) Title  
    - `$CurrentUserStreetaddress$`: Street address  
    - `$CurrentUserPostalcode$`: Postal code  
    - `$CurrentUserLocation$`: Location  
    - `$CurrentUserState$`: State  
    - `$CurrentUserCountry$`: Country  
    - `$CurrentUserTelephone$`: Telephone number  
    - `$CurrentUserFax$`: Facsimile number  
    - `$CurrentUserMobile$`: Mobile phone  
    - `$CurrentUserMail$`: E-mail address  
    - `$CurrentUserPhoto$`: Photo from Active Directory, see "[12.1 Photos from Active Directory](#121-photos-from-active-directory)" for details  
    - `$CurrentUserPhotodeleteempty$`: Photo from Active Directory, see "[12.1 Photos from Active Directory](#121-photos-from-active-directory)" for details  
    - `$CurrentUserExtattr1$` to `$CurrentUserExtattr15$`: Exchange extension attributes 1 to 15  
    - `$CurrentUserOffice$`: Office room number (physicalDeliveryOfficeName)  
    - `$CurrentUserCompany$`: Company  
    - `$CurrentUserMailnickname$`: Alias (mailNickname)  
    - `$CurrentUserDisplayname$`: Display Name  
- Manager of currently logged-in user  
    - Same variables as logged-in user, `$CurrentUserManager[...]$` instead of `$CurrentUser[...]$`  
- Current mailbox  
    - Same variables as logged-in user, `$CurrentMailbox[...]$` instead of `$CurrentUser[...]$`  
- Manager of current mailbox  
    - Same variables as logged-in user, `$CurrentMailboxManager[...]$` instead of `$CurrentMailbox[...]$`  
## 13.1. Photos from Active Directory  
The script supports replacing images in signature templates with photos stored in Active Directory.

When using images in OOF templates, please be aware that Exchange and Outlook do not yet support images in OOF messages.

As with other variables, photos can be obtained from the currently logged-in user, it's manager, the currently processed mailbox and it's manager.
  
### 13.1.1. When using DOCX template files
To be able to apply Word image features such as sizing, cropping, frames, 3D effects etc, you have to exactly follow these steps:  
1. Create a sample image file which will later be used as placeholder.  
2. Optionally: If the sample image file name contains one of the following variable names, the script recognizes it and you do not need to add the value to the alternative text of the image in step 4:  
    - `$CurrentUserPhoto$`  
    - `$CurrentUserPhotodeleteempty$`  
    - `$CurrentUserManagerPhoto$`  
    - `$CurrentUserManagerPhotodeleteempty$`  
    - `$CurrentMailboxPhoto$`  
    - `$CurrentMailboxPhotodeleteempty$`  
    - `$CurrentMailboxManagerPhoto$`  
    - `$CurrentMailboxManagerPhotodeleteempty$`  
3. Insert the image into the signature template. Make sure to use `Insert | Pictures | This device` (Word 2019, other versions have the same feature in different menus) and to select the option `Insert and Link` - if you forget this step, a specific Word property is not set and the script will not be able to replace the image.  
4. If you did not follow optional step 2, please add one of the following variable names to the alternative text of the image in Word (these variables are removed from the alternative text in the final signature):  
    - `$CurrentUserPhoto$`  
    - `$CurrentUserPhotodeleteempty$`  
    - `$CurrentUserManagerPhoto$`  
    - `$CurrentUserManagerPhotodeleteempty$`  
    - `$CurrentMailboxPhoto$`  
    - `$CurrentMailboxPhotodeleteempty$`  
    - `$CurrentMailboxManagerPhoto$`  
    - `$CurrentMailboxManagerPhotodeleteempty$`  
5. Format the image as wanted.

For the script to recognize images to replace, you need to follow at least one of the steps 2 and 4. If you follow both, the script first checks for step 2 first. If you provide multiple image replacement variables, `$CurrentUser[...]$` has the highest priority, followed by `$CurrentUserManager[...]$`, `$CurrentMailbox[...]$` and `$CurrentMailboxManager[...]$`. It is recommended to use only one image replacement variable per image.  
  
The script will replace all images meeting the conditions described in the steps above and replace them with Active Directory photos in the background. This keeps Word image formatting option alive, just as if you would use Word's `"Change picture"` function.  

### 13.1.2. When using HTM template files
Images are replaced when the `src` or `alt` property of the image tag contains one of the following strings:
- `$CurrentUserPhoto$`  
- `$CurrentUserPhotodeleteempty$`  
- `$CurrentUserManagerPhoto$`  
- `$CurrentUserManagerPhotodeleteempty$`  
- `$CurrentMailboxPhoto$`  
- `$CurrentMailboxPhotodeleteempty$`  
- `$CurrentMailboxManagerPhoto$`  
- `$CurrentMailboxManagerPhotodeleteempty$`

Be aware that Outlook does not support the full HTML feature set. For example:
- Some (older) Outlook versions ignore the `width` and `height` properties for embedded images.  
  To overcome this limitation, use images in a connected folder (such as `Test all default replacement variables.files` in the sample templates folder) and additionally set the Set-OutlookSignatures parameter `EmbedImagesInHtml` to ``false`.
- Text and image formatting are limited, especially when HTML5 or CSS features are used.
- Consider switching to DOCX templates for easier maintenance.
### 13.1.3. Common behavior
If there is no photo available in Active Directory, there are two options:  
- You used the `$Current[...]Photo$` variables: The sample image used as placeholder is shown in the signature.  
- You used the `$Current[...]PhotoDeleteempty$` variables: The sample image used as placeholder is deleted from the signature, which may affect the layout of the remaining signature depending on your formatting options.

**Attention**: A signature with embedded images has the expected file size in DOCX, HTML and TXT formats, but the RTF file will be much bigger.

The signature template `.\templates\Signatures DOCX\Test all signature replacement variables.docx` contains several embedded images and can be used for a file comparison:  
- .docx: 23 KB  
- .htm: 87 KB  
- .RTF without workaround: 27.5 MB  
- .RTF with workaround: 1.4 MB
  
The script uses a workaround, but the resulting RTF files are still huge compared to other file types and especially for use in emails. If this is a problem, please either do not use embedded images in the signature template (including photos from Active Directory), or switch to HTML formatted emails.

If you ran into this problem outside this script, please consider modifying the ExportPictureWithMetafile setting as described in  <a href="https://support.microsoft.com/kb/224663" target="_blank">this Microsoft article</a>.  
If the link is not working, please visit the <a href="https://web.archive.org/web/20180827213151/https://support.microsoft.com/en-us/help/224663/document-file-size-increases-with-emf-png-gif-or-jpeg-graphics-in-word" target="_blank">Internet Archive Wayback Machine's snapshot of Microsoft's article</a>.  
## 13.2. Delete images when attribute is empty, variable content based on group membership
You can avoid creating multiple templates which only differ by the images contained by only creating one template containing all images and marking this images to be deleted when a certain replacement variable is empty.

Just add the text `$<name of the replacement variable>DELETEEMPTY$` (for example: `$CurrentMailboxExtattr10deleteempty$` ) to the description or alt text of the image. Taking the example, the image is deleted when extension attribute 10 of the current mailbox is empty.

This can be combined with the `GroupsSIDs` attribute of the current mailbox or current user to only keep images when the mailbox is member of a certain group.

Examples:
- A signature should only show a social network icon with an associated link when there is data in the extension attribute 10 of the mailbox:
  - Insert the icon of the social network in the template, set the hyperlink target to '$CurrentMailboxExtattr10$' and add '$CurrentMailboxExtattr10Deleteempty$' to the description of the picture.
    - When using embedded and linked pictures, you can also set the file name to '$CurrentMailboxExtattr10Deleteempty$'
- A signature should only contain a certain image when the current mailbox is a member of the Marketing group:
  - Create a new replacement variable. We use '$CurrentMailbox-ismemberof-marketing$' in the following example.
    - Attention on-prem users: If Domain Local Active Directory groups are involved, you need to set the `IncludeMailboxForestDomainLocalGroups` parameter to `true` when running Set-OutlookSignatures, so that the SIDs of these groups are considered too.
    - If the current mailbox is a member, give '$CurrentMailbox-ismemberof-marketing$' any value. If not, give '$CurrentMailbox-ismemberof-marketing$' no value (NULL or an empty string).
    - The code for all this is just one line - it is long, but you only have to modify three strings right at the beginning:
      ```
      # Check if current mailbox is member of group 'EXAMPLEDOMAIN\Marketing' and set $ReplaceHash['$CurrentMailbox-ismemberof-marketing$'] accordingly
      #
      # Replace 'EXAMPLEDOMAIN Marketing' with the domain and group you are searching for. Use 'EntraID' or 'AzureAD' instead of 'EXAMPLEDOMAIN' to only search Entra ID/Azure AD/Graph
      # Replace '$CurrentMailbox-ismemberof-marketing$' with the replacement variable that should be used
      # Replace 'CurrentMailbox' with 'CurrentUser' if you do not want to check the current mailbox group SIDs, but the group SIDs of the current user's mailbox
      #
      # The 'GroupsSIDs' attribute is available for the current mailbox and the current user, but not for the managers of these two
      #   It contains the mailboxes' SID and SIDHistory, the SID and SIDHistory of all groups the mailbox belongs to (nested), and also considers group membership (nested) across trusts.
      #   Attention on-prem users: If Active Directory groups of the Domain Local type are queried, you need to set the `IncludeMailboxForestDomainLocalGroups` parameter to `true` when running Set-OutlookSignatures, so that the SIDs of these groups are considered in GroupsSIDs, too.
      #
      @(@('CurrentMailbox', '$CurrentMailbox-IsMemberOf-Marketing$', 'EXAMPLEDOMAIN Marketing'), @()) | Where-Object { $_ } | Foreach-Object { if ((Get-Variable -Name "ADProps$($_[0])" -ValueOnly).GroupsSids -icontains ResolveToSid($_2]) ) { $ReplaceHash[$_[1]] = 'yes' } else { $ReplaceHash[$_[1]] = $null } }
      ```
  - Insert the image in the template, and add '$CurrentMailbox-IsMemberOf-MarketingDeleteempty$' to the description of the picture.
    - When using embedded and linked pictures, you can also set the file name to '$CurrentMailbox-IsMemberOf-MarketingDeleteempty$'

# 14. Outlook Web  
If the currently logged-in user has configured his personal mailbox in Outlook, the default signature for new emails is configured in Outlook Web automatically.

If the default signature for new mails matches the one used for replies and forwarded e-mail, this is also set in Outlook.

If different signatures for new and reply/forward are set, only the new signature is copied to Outlook Web.

If only a default signature for replies and forwards is set, only this new signature is copied to Outlook Web.

If there is no default signature in Outlook, Outlook Web settings are not changed.

All this happens with the credentials of the currently logged-in user, without any interaction neccessary.  
# 15. Hybrid and cloud-only support
Set-OutlookSignatures supports three directory environments:
- Active Directory on premises. This requires direct connection to Active Directory Domain Controllers, which usually only works when you are connected to your company network.
- Hybrid. This environment consists of an Active Directory on premises, which is synced with Microsoft 365 Azure Active Directory in the cloud. Make sure that all signature relevant groups (if applicable) are available as well on-prem and in the cloud, and also ensure this for mail related attributes (At least legacyExchangeDN, msexchrecipienttypedetails, msExchMailboxGuid and proxyaddresses - see https://learn.microsoft.com/en-us/azure/active-directory/hybrid/connect/reference-connect-sync-attributes-synchronized for details. Make sure that the mail attribute in any environment is set to the users primary SMTP address - it may only be empty on the linked user account in the on-prem resource forest scenario.). If the script can't make a connection to your on-prem environment, it tries to get required data from the cloud via the Microsoft Graph API.
- Cloud-only. This environment has no Active Directory on premises, or does not sync mail attributes between the cloud and the on-prem enviroment. The script does not connect to your on-prem environment, only to the cloud via the Microsoft Graph API.

The script parameter `GraphOnly` defines which directory environment is used:
- `-GraphOnly false` or not passing the parameter: On-prem AD first, Entra ID/Azure AD only when on-prem AD can not be reached
- `-GraphOnly true`: Entra ID/Azure AD only, even when on-prem AD could be reached
## 15.1. Basic Configuration
To allow communication between Microsoft Graph and Set-Outlooksignatures, both need to be configured for each other.

The easiest way is to once start Set-OutlookSignatures with a cloud administrator. The administrator then gets asked for admin consent for the correct permissions:
1. Log on with a user that has administrative rights in Entra ID/Azure AD.
2. Run `Set-OutlookSignatures.ps1 -GraphOnly true`
3. When asked for credentials, provide your Entra ID/Azure AD admin credentials
4. For the required permissions, grant consent in the name of your organization

If you don't want to use custom Graph attributes or other advanced configurations, no more configuration in Microsoft Graph or Set-OutlookSignatures is required.

If you prefer using own application IDs or need advanced configuration, follow these steps:  
- In Microsoft Graph, with an administrative account:
  - Create an application with a Client ID
  - Provide admin consent (pre-approval) for the following scopes (permissions):
    - `https://graph.microsoft.com/openid` for logging-on the user
    - `https://graph.microsoft.com/email` for reading the logged-in user's mailbox properties
    - `https://graph.microsoft.com/profile` for reading the logged-in user's properties
    - `https://graph.microsoft.com/user.read.all` for reading properties of other users (manager, additional mailboxes and their managers)
    - `https://graph.microsoft.com/group.read.all` for reading properties of all groups, required for templates restricted to groups
    - `https://graph.microsoft.com/mailboxsettings.readwrite` for updating the logged-in user's Out of Office auto reply messages
    - `https://graph.microsoft.com/EWS.AccessAsUser.All` for updating the logged-in user's Outlook Web signature
  - Set the Redirect URI to `http://localhost` and configure it for `mobile and desktop applications`
  - Enable `Allow public client flows` to make Windows Integrated Authentication (SSO) work for Entra ID/Azure AD joined devices
- In Set-OutlookSignature, use `.\config\default graph config.ps1` as a template for a custom Graph configuration file
  - Set `$GraphClientID` to the application ID created by the Graph administrator before
  - Use the `GraphConfigFile` parameter to make the tool use the newly created Graph configuration file.
## 15.2. Advanced Configuration
The Graph configuration file allows for additional, advanced configuration:
- `$GraphEndpointVersion`: The version of the Graph REST API to use
- `$GraphUserProperties`: The properties to load for each graph user/mailbox. You can add custom attributes here.
- `$GraphUserAttributeMapping`: Graph and Active Directory attributes are not named identically. Set-OutlookSignatures therefore uses a "virtual" account. Use this hashtable to define which Graph attribute name is assigned to which attribute of the virtual account.  
The virtual account is accessible as `$ADPropsCurrentUser[...]` in `.\config\default replacement variables.ps1`, and therefore has a direct impact on replacement variables.
## 15.3. Authentication
In hybrid and cloud-only scenarios, Set-OutlookSignatures automatically tries three stages of authentication.
1. Windows Integrated Authentication  
  This works in hybrid scenarios. The credentials of the currently logged-in user are used to access Microsoft Graph without any further user interaction.
2. Silent authentication  
  If Windows Integrated Authentication fails, the User Principal Name of the currently logged-in user is determined. If an existing cached cloud credential for this UPN is found, it is used for authentication with Microsoft Graph.  
  A default browser window with an "Authentication successful" message may open, it can be closed anytime.
3. User interaction  
  If the other authentication methods fail, the user is interactively asked for credentials. No custom components are used, only the official Microsoft 365 authentication site and the user's default browser. 
# 16. Simulation mode  
Simulation mode is enabled when the parameter `SimulateUser` is passed to the script. It answers the question `"What will the signatures look like for user A, when Outlook is configured for the mailboxes X, Y and Z?"`.

Simulation mode is useful for content creators and admins, as it allows to simulate the behavior of the script and to inspect the resulting signature files before going live.
  
In simulation mode, Outlook registry entries are not considered and nothing is changed in Outlook and Outlook web.

The template files are handled just as during a real script run, but only saved to the folder passed by the parameters AdditionalSignaturePath and AdditionalSignaturePath folder.
  
`SimulateUser` is a mandatory parameter for simulation mode. This value replaces the currently logged-in user. Use a logon name in the format 'Domain\User' or a Universal Principal Name (UPN, looks like an e-mail address, but is not neecessarily one).

`SimulateMailboxes` is optional for simulation mode, although highly recommended. It is a comma separated list of e-mail addresses replacing the list of mailboxes otherwise gathered from the registry.

`SimulateTime` is optional for simulation mode. Simulating a certain time is helpful when time-based templates are used.

**Attention**: Simulation mode only works when the user starting the simulation is at least from the same Active Directory forest as the user defined in SimulateUser.  Users from other forests will not work.  
# 17. FAQ
## 17.1. Where can I find the changelog?
The changelog is located in the `.\docs` folder, along with other documents related to Set-OutlookSignatures.
## 17.2. How can I contribute, propose a new feature or file a bug?
If you have an idea for a new feature or have found a problem, please <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank">create an issue on GitHub</a>.

If you want to contribute code, please have a look at `.\docs\CONTRIBUTING` for a rough overview of the proposed process.
## 17.3. How is the account of a mailbox identified?
The legacyExchangeDN attribute is the preferred method to find the account of a mailbox, as this also works in specific scenarios where the mail and proxyAddresses attribute is not sufficient:
- Separate Active Directory forests for users and Exchange mailboxes: In this case, the mail attribute is usually set in the user forest, although there are no mailboxes in this forest.
- One common e-mail domain across multiple Exchange organizations: In this case, the address book is very like synchronized between Active Directory forests by using contacts or mail-enabled users, which both will have the SMTP address of the mailbox in the proxyAddresses attribute.

The legacyExchangeDN search considers migration scenarios where the original legacyExchangeDN is only available as X500 address in the proxyAddresses attribute of the migrated mailbox, or where the the mailbox in the source system has been converted to a mail enabled user still having the old legacyExchangeDN attribute.

If Outlook does not have information about the legacyExchangeDN of a mailbox (for example, when accessing a mailbox via protocols such as POP3 or IMAP4), the account behind a mailbox is searched by checking if the e-mail address of the mailbox can be found in the proxyAddresses attribute of an account in Active Directory/Graph.

If the account behind a mailbox is found, group membership information be retrieved and group specific templates can be applied.
If the account behind a mailbox is not found, group membership can not be retrieved and group specific templates can not be applied. Such mailboxes can still receive common and mailbox specific signatures and OOF messages.  
## 17.4. How is the personal mailbox of the currently logged-in user identified?  
The personal mailbox of the currently logged-in user is preferred to other mailboxes, as it receives signatures first and is the only mailbox where the Outlook Web signature can be set.

The personal mailbox is found by simply checking if the Active Directory mail attribute of the currently logged-in user matches an SMTP address of one of the mailboxes connected in Outlook.

If the mail attribute is not set, the currently logged-in user's objectSID is compared with all the mailboxes' msExchMasterAccountSID. If there is exactly one match, this mailbox is used as primary one.
  
Please consider the following caveats regarding the mail attribute:  
- When Active Directory attributes are directly modified to create or modify users and mailboxes (instead of using Exchange Admin Center or Exchange Management Shell), the mail attribute is often not updated and does not match the primary SMTP address of a mailbox. Microsoft strongly recommends that the mail attribute matches the primary SMTP address.  
- When using linked mailboxes, the mail attribute of the linked account is often not set or synced back from the Exchange resource forest. Technically, this is not necessary. From an organizational point of view it makes sense, as this can be used to determine if a specific user has a linked mailbox in another forest, and as some applications (such as "scan to e-mail") may need this attribute anyhow.  
## 17.5. Which ports are required?  
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
## 17.6. Why is Out of Office abbreviated OOF and not OOO?  
Back in the 1980s, Microsoft had a UNIX OS named Xenix ... but read yourself <a href="https://techcommunity.microsoft.com/t5/exchange-team-blog/why-is-oof-an-oof-and-not-an-ooo/ba-p/610191" target="_blank">here</a>.  
## 17.7. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.  
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
- `.\templates\Out of Office DOCX` and `.\templates\signatures DOCX` contain templates in the DOCX format  
- `.\templates\Out of Office HTML` contains templates in the HTML format as Word exports them when using `"Website, filtered"` as format. Note the additional folders for each signature.  
- `.\templates\Signatures HTML` contains templates in the HTML format. Note that there are no additional folders, as the Word export files have been processed with ConvertTo-SingleFileHTML function to create a single HTMl file with all local images embedded.  
## 17.8. How can I log the script output?  
The script has no built-in logging option other than writing output to the host window.

You can, for example, use PowerShell's `Start-Transcript` and `Stop-Transcript` commands to create a logging wrapper around Set-OutlookSignatures.ps1.  
## 17.9. How can I get more script output for troubleshooting?
Start the script with the '-verbose' parameter to get the maximum output for troubleshooting.
## 17.10. Can multiple script instances run in parallel?  
The script is designed for being run in multiple instances at the same. You can combine any of the following scenarios:  
- One user runs multiple instances of the script in parallel  
- One user runs multiple instances of the script in simulation mode in parallel  
- Multiple users on the same machine (e.g. Terminal Server) run multiple instances of the script in parallel  

Please see `.\sample code\SimulateAndDeploy.ps1` for an example how to run multiple instances of Set-OutlookSignatures in parallel in a controlled manner. Don't forget to adopt path names and variables to your environment.
## 17.11. How do I start the script from the command line or a scheduled task?  
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
  
If you provided your users a link so they can start Set-OutlookSignatures.ps1 with the correct parameters on their own, you may want to use the official icon: `.\logo\Set-OutlookSignatures Icon.ico`

Please see `.\sample code\Set-OutlookSignatures.cmd` for an example. Don't forget to adopt path names to your environment.
### 17.11.1. Start Set-OutlookSignatures in hidden/invisible mode
Even when the `hidden` parameter is passed to PowerShell, a window is created and minimized. Even when this only takes some tenths of a second, it is not only optically disturbing, but the new window may also steal the keyboard focus.

The only workaround is to start PowerShell from another program, which does not need an own console window. If you do not want to write such a program yourself, you can use the Windows Script Host for this.

Create a .vbs (Visual Basic Script) file, paste and adapt the following code into it:
```
command = "PowerShell.exe -Command ""& '\\server\share\directory\Set-OutlookSignatures.ps1' -SignatureTemplatePath '\\server\share\directory\templates\Signatures DOCX' -OOFTemplatePath '\\server\share\directory\templates\Out of Office DOCX' -ReplacementVariableConfigFile '\\server\share\directory\config\default replacement variables.ps1'"" "

set shell = CreateObject("WScript.Shell")

shell.Run command, 0
'0  Hides the window and activates another window.
'1  Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
'2  Activates the window and displays it as a minimized window.
'3  Activates the window and displays it as a maximized window.
'4  Displays a window in its most recent size and position. The active window remains active.
'5  Activates the window and displays it in its current size and position.
'6  Minimizes the specified window and activates the next top-level window in the Z order.
'7  Displays the window as a minimized window. The active window remains active.
'8  Displays the window in its current state. The active window remains active.
'9  Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
'10 Sets the show state based on the state of the program that started the application.
```

Then, run the .vbs file directly, without specifying cscript.exe as host (just execute `start.vbs`, not `cscript.exe start.vbs`).
## 17.12. How to create a shortcut to the script with parameters?  
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

See `.\sample code\CreateDesktopIcon.ps1` for a code example. Don't forget to adopt path names to your environment. 
## 17.13. What is the recommended approach for implementing the software?  
The Quick Start Guide in this document is a good overall starting point for beginners.

For the organizational aspects around Set-OutlookSignatures, Benefactor Circle members have access to the "Implementation Approach" document. The content is based on real life experiences implementing the script in multi-client environments with a five-digit number of mailboxes.

It contains proven procedures and recommendations for product managers, architects, operations managers, account managers and e-mail and client administrators. It is suited for service providers as well as for clients.

It covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The document is available in English and German language.  
## 17.14. What is the recommended approach for custom configuration files?
You should not change the default configuration files `.\config\default replacement variable.ps1` and `.\config\default graph config.ps1`, as they might be changed in a future release of Set-OutlookSignatures. In this case, you would have to sort out the changes yourself.

The following steps are recommended:
1. Create a new custom configuration file in a separate folder.
2. The first step in the new custom configuration file should be to load the default configuration file, `.\config\default replacement variable.ps1` in this example:
   ```
   # Loading default replacement variables shipped with Set-OutlookSignatures
   . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath '\\server\share\folder\Set-OutlookSignatures\config\default replacement variables.ps1' -Raw)))
   ```
3. After importing the default configuration file, existing replacement variables can be altered with custom definitions and new replacement variables can be added.
4. Start Set-OutlookSignatures with the parameter `ReplacementVariableConfigFile` pointing to the new custom configuration file.
## 17.15. Isn't a plural noun in the script name against PowerShell best practices?
Absolutely. PowerShell best practices recommend using singular nouns, but Set-OutlookSignatures contains a plural noun.

I intentionally decided not to follow the singular noun convention, as another language as PowerShell was initially used for coding and the name of the tool was already defined. If this was a commercial enterprise project, marketing would have overruled development.
## 17.16. The script hangs at HTM/RTF export, Word shows a security warning!?
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
## 17.17. How to avoid blank lines when replacement variables return an empty string?
Not all users have values for all attributes, e. g. a mobile number. These empty attributes can lead to blank lines in signatures, which may not look nice.

Follow these steps to avoid blank lines:
1. Use a custom replacement variable config file.
2. Modify the value of all attributes that should not leave an blank line when there is no text to show:
    - When the attribute is empty, return an empty string
    - Else, return a newline (`Shift+Enter` in Word, `` `n `` in PowerShell, `<br>` in HTML) or a paragraph mark (`Enter` in Word, `` `r`n `` in PowerShell, `<p>` in HTML), and then the attribute value.  
3. Place all required replacement variables on a single line, without a space between them. The replacement variables themselves contain the required newline or paragraph marks.
4. Use the ReplacementVariableConfigFile parameter when running the script.

Be aware that text replacement also happens in hyperlinks (`tel:`, `mailto:` etc.).  
Instead of altering existing replacement variables, it is recommended to create new replacement variables with modified content.  
Use the new one for the pure textual replacement (including the newline), and the original one for the replacement within the hyperlink.  

The following example describes optional preceeding text combined with an optional replacement variable containing a hyperlink.  
The internal variable `$UseHtmTemplates` is used to automatically differentiate between DOCX and HTM line breaks.
- Custom replacement variable config file
  ```
  $ReplaceHash['$CurrentUserTelephone-prefix-noempty$'] = $(if (-not $ReplaceHash['$CurrentUserTelephone$']) { '' } else { $(if ($UseHtmTemplates) { '<br>' } else { "`n" }) + 'Telephone: ' } )
  $ReplaceHash['$CurrentUserMobile-prefix-noempty$'] = $(if (-not $ReplaceHash['$CurrentUserMobile$']) { '' } else { $(if ($UseHtmTemplates) { '<br>' } else { "`n" }) + 'Mobile: ' } )
  ```
- Word template:  
  <pre><code>E-Mail: <a href="mailto:$CurrentUserMail$">$CurrentUserMail$</a>$CurrentUserTelephone-prefix-noempty$<a href="tel:$CurrentUserTelephone$">$CurrentUserTelephone$</a>$CurrentUserMobile-prefix-noempty$<a href="tel:$CurrentUserMobile$">$CurrentUserMobile$</a></code></pre>

  Note that all variables are written on one line and that not only `$CurrentUserMail$` is configured with a hyperlink, but `$CurrentUserPhone$` and `$CurrentUserMobile$` too:
  - `mailto:$CurrentUserMail$`
  - `tel:$CurrentUserTelephone$`
  - `tel:$CurrentUserMobile$`
- Results
  - Telephone number and mobile number are set.  
  The paragraph marks come from `$CurrentUserTelephone-prefix-noempty$` and `$CurrentUserMobile-prefix-noempty$`.  
    <pre><code>E-Mail: <a href="mailto:first.last@example.com">first.last@example.com</a>
    Telephone: <a href="tel:+43xxx">+43xxx</a>
    Mobile: <a href="tel:+43yyy">+43yyy</a></code></pre>
  - Telephone number is set, mobile number is empty.  
  The paragraph mark comes from `$CurrentUserTelephone-prefix-noempty$`.  
    <pre><code>E-Mail: <a href="mailto:first.last@example.com">first.last@example.com</a>
    Telephone: <a href="tel:+43xxx">+43xxx</a></code></pre>
  - Telephone number is empty, mobile number is set.  
  The paragraph mark comes from `$CurrentUserMobile-prefix-noempty$`.  
    <pre><code>E-Mail: <a href="mailto:first.last@example.com">first.last@example.com</a>
    Mobile: <a href="tel:+43yyy">+43yyy</a></code></pre>
## 17.18. Is there a roadmap for future versions?
There is no binding roadmap for future versions, although I maintain a list of ideas in the 'Contribution opportunities' chapter of '.\docs\CONTRIBUTING'.

Fixing issues has priority over new features, of course.
## 17.19. How to deploy signatures for "Send As", "Send On Behalf" etc.?
The script only considers primary mailboxes, these are mailboxes added as separate accounts. This is the same way Outlook handles mailboxes from a signature perspective: Outlook can not handle signatures for non-primary mailboxes (added via "Open these additional mailboxes").

If you want to deploy signatures for non-primary mailboxes, sett the parameter `SignaturesForAutomappedAndAdditionalMailboxes` to `true` to allow the script to detect automapped and additional mailboxes. Signatures can be deployed for these types of mailboxes, but they can not be set as default signatures due to technical restrictions in Outlook.

If you want to deploy signatures for
- mailboxes you don't add to Outlook but just use an assigned "Send As" or "Send on Behalf" right by choosing a different "From" address,
- distribution lists, for which you use an assigned "Send As" or "Send on Behalf" right by choosing a different "From" address,
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
## 17.20. Can I centrally manage and deploy Outook stationery with this script?
Outlook stationery describes the layout of e-mails, including font size and color for new e-mails and for replies and forwards.

The default e-mail font, size and color are usually an integral part of corporate design and corporate identity. CI/CD typically also defines the content and layout of signatures.

Set-OutlookSignatures has no features regarding deploying Outlook stationery, as there are better ways for doing this.  
Outlook stores stationery settings in `HKCU\Software\Microsoft\Office\<Version>\Common\MailSettings`. You can use a logon script or group policies to deploy these keys, on-prem and for managed devices in the cloud.  
Unfortunately, Microsoft's group policy templates (ADMX files) for Office do not seem to provide detailed settings for Outlook stationery, so you will have to deploy registry keys. 
## 17.21. Why is dynamic group membership not considered on premises?
Membership in dynamic groups, no matter if they are of the security or distribution type, is considered only when using Microsoft Graph.

Dynamic group membership is not considered when using an on premises Active Directory. 

The reason for this is that Graph and on-prem AD handle dynamic group membership differently:
### 17.21.1. Graph
Microsoft Graph caches information about dynamic group membership at the group as well as at the user level.  

Graph regularly executes the LDAP queries defining dynamic groups and updates existing attributes with member information.  
Dynamic groups in Graph are therefore not strictly dynamic in terms of running the defining LDAP query every time a dynamic group is used and thus providing near real-time member information - they behave more like regularly updated static groups, which makes handling for scripts and applications much easier.

For the usecases of Set-OutlookSignatures, there is no difference between a static and a dynamic group in Graph:
- Querying the `transitiveMemberOf` attribute of a user returns static as well as dynamic group membership.
- Querying the `members` attribute of a group returns the group's members, no matter if the group is static or dynamic.
### 17.21.2. Active Directory on premises
Active Directory on premises does not cache any information about membership in dynamic groups at the user level, so dynamic groups do not appear in attributes such as `memberOf` and `tokenGroups`.  
Active Directory on premises also does not cache any information about members of dynamic groups at the group level, so the group attribute `members` is always empty. 

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
## 17.22. Why is no admin or user GUI available?
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

## 17.23. What if a user has no Outlook profile or is prohibited from starting Outlook?
If a user has never started Outlook before or has deleted all Outlook profiles, Set-OutlookSignatures will still be useful: It will create the signature folder if it does not exist, determine the logged-in users e-mail address, create the signatures for his personal mailbox, set a default signature in Outlook Web as well as the Out of Office messages.

Default signatures can not be set locally or in Outlook Web until an Outlook profile has been configured, as the corresponding settings are stored in registry paths containing random numbers, which need to be created by Outlook.
## 17.24. What if Outlook is not installed at all?
If Outlook is not installed at all, Set-OutlookSignatures will still be useful: It determine the logged-in users e-mail address, create the signatures for his personal mailbox in a temporary location, set a default signature in Outlook Web as well as the Out of Office messages.
## 17.25. What about the roaming signatures feature in Exchange Online?  
Microsoft changed how and where signatures are stored for mailboxes in Exchange Online. Basically, signatures are stored in the mailbox itself.  
For details, please see <a href="https://support.microsoft.com/en-us/office/outlook-roaming-signatures-420c2995-1f57-4291-9004-8f6f97c54d15?ui=en-us&rs=en-us&ad=us" target="_blank">this Microsoft article</a>.  

This is a good idea, as it makes signatures available across devices and apps.

Before we jump into the details of roaming signatures and the pros and cons of them: **Set-OutlookSignatures can experimentally handle roaming signatures since v4.0.0!** See `MirrorLocalSignaturesToCloud` in this document for details.

Some personal educated guesses based on available documentation, Outlook for Windows beta versions and several Exchange Online tenants:
- The feature has first been annount by Microsoft in 2020, but has been postponed multiple times. The feature seems to get enabled in waves across all Exchange Online tenants since late 2022.
- Microsoft has not yet published a public API, although the feature is announced since 2020 and being actively enabled. 
- Outlook for Windows is the only client mentioned to support the new feature for now. I am confident more e-mail clients - especially Outlook for Mac, iOS and Android - will follow (the sooner, the better).
- Although Micrsoft actively enables the feature for all Exchange Online tenants, it is yet unclear if this feature is available for shared mailboxes. If yes, the disadvantage is that signatures for shared mailboxes can no longer be personalized, as the latest signature change would be propagated to all users accessing the shared mailbox (which is especially bad when personalized signatures for shared mailboxes are set as default signature - think about `$CurrentUser[...]$` replacement variables).
- The roaming signatures feature is only available for mailboxes in the cloud. Mailboxes on on-prem servers do not (yet?) support this feature, no matter if in pure on-prem or in hybrid scenarios.
- Outlook for Windows (beta) versions already support the roaming signatures feature. Until an API is available, you can disable the feature with a registry key. This forces Outlook for Windows to use the well-known file based approach and ensure full compatibility with Set-OutlookSignatures, until a public API is released and incorporated into the script. For details, please see <a href="https://support.microsoft.com/en-us/office/outlook-roaming-signatures-420c2995-1f57-4291-9004-8f6f97c54d15?ui=en-us&rs=en-us&ad=us" target="_blank">this Microsoft article</a>.
  - With the `DisableRoamingSignatures` (formerly named `DisableRoamingSignaturesTemporaryToggle`) registry value being absent or set to 0, file based signatures created by tools such as Set-OutlookSignatures are regularly deleted and replaced with signatures stored directly in the mailbox.
  - With the `DisableRoamingSignatures` (formerly named `DisableRoamingSignaturesTemporaryToggle`) registry value set to 1, the file based approach continues to work as known. Outlook does not synchronize signatures to the mailbox.

Microsoft is already supporting the feature in Outlook Web for more and more Exchange Online tenants. Currently, this breaks PowerShell commands such as Set-MailboxMessageConfiguration and there is no public API available.
  - Set-OutlookSignatures can set one Outlook Web signature, but an Exchange Online tenant with multiple signatures feature enabled just ignores this signature (see the next chapter for workarounds).

If you do not want to use Group Policy, you can use the following PowerShell script to disable roaming signatures in Outlook:
```
# Define variables
$RegistryPath = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup'
$RegistryName = 'DisableRoamingSignatures' # formerly named 'DisableRoamingSignaturesTemporaryToggle'
$RegistryType = 'DWORD'
$RegistryValue = '1'

# Create the path if it does not exist
If (-not (Test-Path $RegistryPath)) {
  $null = New-Item -Path $RegistryPath -Force
}  

# Create/update the name-value-pair
$null = New-ItemProperty -Path $RegistryPath -Name $RegistryName -Value $RegistryValue -PropertyType $RegistryType -Force

# Query the name-value-pair
Get-ItemProperty -Path $RegistryPath -Name $RegistryName
```
### 17.25.1. Please be aware of the following problem
Since Q3 2021, the roaming signature feature appears and disappears on Outlook Web of cloud mailboxes. There is still no hint of an API, or a way to disable it on your own in Exchange Online.

When multiple signatures in Outlook Web are enabled, Set-OutlookSignatures can successfully set the signature in Outlook Web, but this signature is ignored.

There is no programmatic way to detect or change this behavior.  
The built-in Exchange Online PowerShell-Cmdlet `Set-MailboxMessageConfiguration` has the same problem, so it seems different Microsoft teams work on different development and release schedules.

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
## 17.26. Why does the text color of my signature change sometimes?
Set-OutlookSignatures does not change text color. Very likely, your template files and your Outlook installation are configured for this color change:
- Per default, Outlook uses black text for new e-mails, and blue text for replies and forwarded e-mails
- Word and the signature editor integrated in Outlook have a specific color named "Automatic"

When using DOCX templates with parts of the text formatted in the "Automatic" color, Outlook changes the color of these parts to black for new e-mails, and to blue for replies and forwards.

This behavior is very often wanted, so that the greeting formula, which usually is part of the signature, has the same color as the preceding text of the e-mail.

The default colors can be configured in Outlook.  
Outlook seems to have problems with this in certain patch levels when creating a reply in the preview pane, popping out the draft to it's own window and then switching to another signature.
## 17.27. How to make Set-OutlookSignatures work with Microsoft Information Protection?
Set-OutlookSignatures does work well with Microsoft Information Protection, when configured correctly.

If you do not enforce setting sensitivity labels or exclude DOCX and RTF file formats, no further actions are required.

If you enforce setting sensitivity labels:
- When using DOCX templates, just set the desired sensitivity label on all your template files.
  - It is recommended to use the 'General' label:
    - Outlook signatures and Out of Office messages usually only contain information which is intended to be shared publicly by design.
    - The templates themselves usually do not contain sensitive data, only placeholder variables.
    - Documents labeled 'General' can be opened without having the Information Protection Add-In for Office installed. This is useful when not all of your Set-OutlookSignatures users are also Information Protection users and have the Add-In installed.
  - When using a template with a sensitivity label other than 'General', every client Set-OutlookSignatures runs on needs the Information Protection Add-In for Office installed, and the user running Set-OutlookSignatures needs permission to access the protected file.
  - The RTF signature file will be created with the same sensitivity label as the template. This is only relevant for the user composing a new e-mail in RTF format, as the composing user needs to be able to open the RTF document and copy the content from it - the actual signature in the e-mail does not have Information Protection applied.
  - The .HTM and .TXT signature files will be created without a sensitivity label, as these documents can not be protected by Microsoft Information Protection.
  - If you do not set a sensitivity label, Word will prompt the user to choose one each time the unlabeled local copy of a template is converted to .htm, .rtf or .txt.
    - The DOCX sample template files that come with Set-OutlookSignatures do not have a sensitivity label set.
- When using HTM templates, no further actions are required.
  - HTM files can not be assigned a sensitivity label, and converting HTM files to RTF is possible even when sensitivity labels are enforced.
  - Converting HTM files to TXT is also no problem, as both file formats can not be assigned a sensitivity label. 

Additional information that might be of interest for your Information Protection configuration:
- Template files are copied to the local temp directory of the user (PowerShell: `[System.IO.Path]::GetTempPath()`) for further use, with a randomly generated GUID as name. The extension is the one of the template (.docx or .htm).
- The local copy of a template file is opened for variable replacement, saved back to disk, and then re-opened for each file conversion (to .htm if neccessary, and optionally to .rtf and/or .txt). 
- Converted files are also stored in the temp directory, using the same GUID as the original file as file name but a different file extension (.htm, .rtf, .txt).
- After all variable replacements and conversions are completed for a template, the converted files (HTM mandatory, RTF and TXT optional) are copied to the Outlook signature folder. The path of this folder is language and version dependent (Registry: `HKCU:\Software\Microsoft\Office\<Outlook Version>\Common\General\Signatures`).
- All temporary files mentioned are deleted by Set-OutlookSignatures as part of the clean-up process.
## 17.28. Images in signatures have a different size than in templates, or a black background
The size of images in signatures may differ to the size of the very same image in a template. This may have observable in several ways:
- Images are already displayed too big or too small when composing a message. Not all signatures with images need to be affected, and the problem does not need to be bound to specific users or client computers.
- Images are displayed correctly when composing and sending an e-mail, but are shown in different sizes at the recipient.

In both cases, usually only e-mails composed HTML format are affected, but not e-mails in RTF format.

When only the recipient is affected, it is very likely that the problem is to be found within the mail client of the recipient, as it very likely does not respect or interpret HTML width and height attributes correctly.  
- This problem can not be solved on the sending side, only on the recipient side. But the sender side can implement a workaround: Do not scale images in templates (by resizing them in Word or using HTML width and height tags), but use the original size of the image. It may be neccessary to resize the images with tools like GIMP before using them in templates.

When the problem can already be seen when composing a message, there may be different root causes and according solutions or workarounds.

To find the root cause:
- Use the same signature template to create individual signatures for all the following steps.
- Find out if the problem is user or computer related. Let affected users log on to non-affected computer, and vice versa, to test this.
- Find out if only Outlook displays the image in the wrong size. Open the signature HTM file in Word, Chrome, Edge and Firefox for comparison.
- Copy the affected HTM signature file (the signature, not the template) and let a non-affected user use it in Outlook to see if the problem exists there, too.
- Compare the 'img' tag between the signature (from the same template) of an affected and a non-affected user. If they are identical, the root cause is not the generated HTML code, but it's interpretation and display in Outlook (therefore, the problem can't be in Set-OutlookSignatures).
- Collect the following data for a number of affected and non-affected users and computer to help you find the root cause:
  - User name
  - Computer name
  - Windows version including build number
  - Word version including build number
  - Outlook version including build number
  - Does Chrome display the image in the correct size?
  - Does Edge display the image in the correct size?
  - Does Firefox display the image in the correct size?
  - Does Outlook display the image in the correct size?
  - Does Word display the image in the correct size?

Two workarounds are available when you do not want to or can't find and solve the root cause of the problem:
- Do not scale images in templates (by resizing them in Word, or using HTML width and height attributes), but use the original size of the image. It may be neccessary to resize the images with tools like GIMP before using them in templates.
- The problem may only appear when templates are converted to signatures on computers configured with a display scaling higher than 100 %. In this case, the problem is in the Word conversion module or the HTML rendering engine of Word (which is used by Outlook). The registry key described in <a href="https://learn.microsoft.com/en-US/outlook/troubleshoot/user-interface/graphics-file-attachment-grows-larger-in-recipient-email" target="_blank">this Microsoft article</a> may help here. After setting the registry key according to the article, Outlook and Word need to be restarted and Set-OutlookSignatures needs to run again.  
Starting with v4.0.0, Set-OutlookSignatures sets the `DontUseScreenDpiOnOpen` registry key to the recommended value. 

Nonetheless, some scaling and display problems simply can not be solved in the HTML code of the signature, because the problem is in the Word HRML rendering engine used by Outlook: For example, some Word builds ignore embedded image width and height attributes and always scale these images at 100% size, or sometimes display them with inverted colors or a black background.  
In this case, you can influence how images are displayed and converted from DOCX to HTM with the parameters `EmbedImagesInHtml` and `DocxHighResImageConversion`:

| Parameter  | Default value | Alternate<br>configuration A | Alternate<br>configuration B | Alternate<br>configuration C |
| :- | :- | :- | :- | :- |
| EmbedImagesInHtml | false | true | true | false |
| DocxHighResImageConversion | true | false | true | false |
| Influence on images | HTM signatures with images consist of multiple files<br><br>Make sure to set the Outlook registry value "Send Pictures With Document" to 1, as described in the documentation of the `EmbedImagesInHtml` parameter. | HTM signatures with images consist of a single file<br><br>Office 2013 can't handle embedded images<br><br>Some versions of Office/Outlook/Word (some Office 2016 builds, for example) show embedded images wrongly sized<br><br>Images can look blurred and pixelated, especially on systems with high display resolution | HTM signatures with images consist of a single file<br><br>Office 2013 can't handle embedded images<br><br>Some versions of Office/Outlook/Word (some Office 2016 builds, for example) show embedded images wrongly sized | HTM signatures with images consist of multiple files<br><br>Images can look blurred and pixelated, especially on systems with high display resolution<br><br>Make sure to set the Outlook registry value "Send Pictures With Document" to 1, as described in the documentation of the EmbedImagesInHtml parameter |
| Recommendation | This configuration should be used as long as there is nothing to the contrary | This configuration should not be used due to the low graphic quality | This configuration may lead to wrongly sized images or images with black background due to a bug in some Office versions | This configuration should not be used due to the low graphic quality |
||||||
